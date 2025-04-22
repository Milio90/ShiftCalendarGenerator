import docx
from icalendar import Calendar, Event
from datetime import datetime, timedelta, date
import re
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import tempfile
import platform
import shutil

def convert_doc_to_docx(doc_path):
    """Convert a .doc file to .docx format using available tools."""
    file_name, file_ext = os.path.splitext(doc_path)
    
    # If already a docx file, return the original path
    if file_ext.lower() == '.docx':
        return doc_path
    
    # Create a temporary output file
    temp_dir = tempfile.gettempdir()
    base_name = os.path.basename(file_name)
    output_path = os.path.join(temp_dir, f"{base_name}_converted.docx")
    
    conversion_successful = False
    error_message = ""
    
    # Try LibreOffice first (cross-platform)
    try:
        # Determine LibreOffice executable based on platform
        libreoffice_cmd = None
        if platform.system() == "Windows":
            # Check common install locations
            possible_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    libreoffice_cmd = path
                    break
        elif platform.system() == "Darwin":  # macOS
            libreoffice_cmd = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        else:  # Linux and others
            libreoffice_cmd = "libreoffice"
        
        if libreoffice_cmd:
            process = subprocess.run([
                libreoffice_cmd,
                "--headless",
                "--convert-to", "docx",
                "--outdir", temp_dir,
                doc_path
            ], capture_output=True, text=True, timeout=30)
            
            # LibreOffice sometimes creates with original filename in the output dir
            expected_file = os.path.join(temp_dir, f"{base_name}.docx")
            if os.path.exists(expected_file):
                # Rename to our expected output path
                shutil.move(expected_file, output_path)
                conversion_successful = True
            else:
                error_message = f"LibreOffice conversion output file not found: {expected_file}"
    except Exception as e:
        error_message = f"LibreOffice conversion failed: {str(e)}"
    
    # If LibreOffice failed, try Microsoft Word automation (Windows only)
    if not conversion_successful and platform.system() == "Windows":
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            doc = word.Documents.Open(doc_path)
            doc.SaveAs(output_path, FileFormat=16)  # 16 = docx format
            doc.Close()
            word.Quit()
            
            if os.path.exists(output_path):
                conversion_successful = True
            else:
                error_message += "\nMicrosoft Word conversion output file not found."
        except Exception as e:
            error_message += f"\nMicrosoft Word conversion failed: {str(e)}"
    
    if conversion_successful:
        print(f"Successfully converted {doc_path} to {output_path}")
        return output_path
    else:
        print(f"Failed to convert .doc to .docx: {error_message}")
        raise Exception(f"Could not convert {doc_path} to .docx format. Please convert it manually and try again.")

def browse_file(title="Select File"):
    """Open a file browser dialog and return the selected file path."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Word Documents", "*.docx *.doc"), ("All Files", "*.*")]
    )
    return file_path

def read_docx_tables(file_path):
    """Read all tables content from a DOCX file."""
    try:
        # Check if file is .doc and convert if needed
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext == '.doc':
            print("Converting .doc file to .docx format...")
            file_path = convert_doc_to_docx(file_path)
        
        doc = docx.Document(file_path)
        if not doc.tables:
            print("No tables found in the document.")
            return []
        
        tables_data = []
        
        for table_index, table in enumerate(doc.tables):
            rows = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                # Skip empty rows
                if any(row_data):
                    rows.append(row_data)
            
            tables_data.append(rows)
            print(f"Table {table_index+1}: Found {len(rows)} rows with data")
        
        return tables_data
    except Exception as e:
        print(f"Error reading document: {e}")
        if "Could not convert" in str(e):
            # This is our custom error from conversion function
            print(str(e))
        messagebox.showerror("Error", f"Could not process the document: {e}")
        return []

def parse_first_table(rows, month, year):
    """Parse the first table format (Regular and On-Call shifts)."""
    shifts = []
    
    for row in rows:
        if len(row) < 4:  # Ensure row has enough columns
            continue
        
        try:
            # Extract day, month, day_of_week, and employees
            day = row[0].strip()
            day_of_week = row[2].strip()
            employees_cell = row[3].strip()
            
            # Skip header rows or rows without day number
            if not day.isdigit():
                continue
            
            day = int(day)
            
            # Parse employee names (may contain two employees, one with asterisk)
            employees = employees_cell.split('\n')
            employees = [e.strip() for e in employees if e.strip()]
            
            for employee in employees:
                is_on_call = "*" in employee
                employee_name = employee.replace("*", "").strip()
                
                # Create shift date
                shift_date = date(year, month, day)
                
                shift_type = "On-Call Shift" if is_on_call else "Regular Shift"
                
                shifts.append({
                    'employee': employee_name,
                    'date': shift_date,
                    'day_of_week': day_of_week,
                    'shift_type': shift_type
                })
        except Exception as e:
            print(f"Error parsing row in first table {row}: {e}")
            continue
    
    return shifts

def parse_second_table(rows, month, year):
    """Parse the second table format (Μεγάλη, Μικρή, ΤΕΠ shifts)."""
    shifts = []
    
    for row in rows:
        if len(row) < 6:  # Ensure row has enough columns for second table format
            continue
        
        try:
            # Extract day, month, day_of_week, and employees from different shifts
            day = row[0].strip()
            day_of_week = row[2].strip()
            megali_shift = row[3].strip()
            mikri_shift = row[4].strip()
            tep_shift = row[5].strip()
            
            # Skip header rows or rows without day number
            if not day.isdigit():
                continue
            
            day = int(day)
            shift_date = date(year, month, day)
            
            # Process Μεγάλη shift (24h)
            if megali_shift:
                employee_name = megali_shift.replace(">", "").strip()
                if employee_name:
                    shifts.append({
                        'employee': employee_name,
                        'date': shift_date,
                        'day_of_week': day_of_week,
                        'shift_type': "Μεγάλη Shift (24h)"
                    })
            
            # Process Μικρή shift (24h)
            if mikri_shift:
                employee_name = mikri_shift.replace(">", "").strip()
                if employee_name:
                    shifts.append({
                        'employee': employee_name,
                        'date': shift_date,
                        'day_of_week': day_of_week,
                        'shift_type': "Μικρή Shift (24h)"
                    })
            
            # Process ΤΕΠ shift (12h)
            if tep_shift:
                employee_name = tep_shift.replace(">", "").strip()
                if employee_name:
                    shifts.append({
                        'employee': employee_name,
                        'date': shift_date,
                        'day_of_week': day_of_week,
                        'shift_type': "TEP Shift (12h)"
                    })
                
        except Exception as e:
            print(f"Error parsing row in second table {row}: {e}")
            continue
    
    return shifts

def parse_specialty_on_call_table(rows):
    """Parse the specialty on-call table format with date (DD-MM-YYYY) in first column."""
    shifts = []
    
    for row in rows:
        if len(row) < 3:  # Ensure row has enough columns
            continue
        
        try:
            # Extract date, day_of_week, and employee
            date_str = row[0].strip()
            day_of_week = row[1].strip()
            employee_name = row[2].strip()
            
            # Skip header rows or rows without proper date format
            if not re.match(r"\d{1,2}-\d{1,2}-\d{4}", date_str):
                continue
            
            # Parse date (DD-MM-YYYY)
            day, month, year = map(int, date_str.split('-'))
            shift_date = date(year, month, day)
            
            if employee_name:
                shifts.append({
                    'employee': employee_name,
                    'date': shift_date,
                    'day_of_week': day_of_week,
                    'shift_type': "On-Call Specialty",  # Will be updated when adding to all_shifts
                })
                
        except Exception as e:
            print(f"Error parsing row in specialty on-call table {row}: {e}")
            continue
    
    return shifts

def create_calendar_for_employee(shifts, employee_name, output_file, cath_lab_shifts=None, ep_shifts=None):
    """Create an iCalendar file with all-day events for a specific employee."""
    # Filter shifts for this specific employee
    employee_shifts = [s for s in shifts if s['employee'].lower() == employee_name.lower()]
    
    # Also check if the employee has any cath lab or EP shifts
    employee_cath_lab_shifts = []
    employee_ep_shifts = []
    
    if cath_lab_shifts:
        employee_cath_lab_shifts = [s for s in cath_lab_shifts if s['employee'].lower() == employee_name.lower()]
        
    if ep_shifts:
        employee_ep_shifts = [s for s in ep_shifts if s['employee'].lower() == employee_name.lower()]
    
    if not employee_shifts and not employee_cath_lab_shifts and not employee_ep_shifts:
        print(f"No shifts found for employee: {employee_name}")
        return None
    
    cal = Calendar()
    cal.add('prodid', '-//Employee Shift Calendar//example.com//')
    cal.add('version', '2.0')
    cal.add('calscale', 'GREGORIAN')
    
    # Group shifts by date to combine multiple shifts on the same day
    shifts_by_date = {}
    
    # Add regular shifts to the grouping
    for shift in employee_shifts:
        date_key = shift['date'].isoformat()
        if date_key not in shifts_by_date:
            shifts_by_date[date_key] = []
        shifts_by_date[date_key].append(shift)
    
    # Add cath lab shifts if they don't overlap with existing dates
    for shift in employee_cath_lab_shifts:
        date_key = shift['date'].isoformat()
        if date_key not in shifts_by_date:
            shifts_by_date[date_key] = []
        shifts_by_date[date_key].append(shift)
    
    # Add EP shifts if they don't overlap with existing dates
    for shift in employee_ep_shifts:
        date_key = shift['date'].isoformat()
        if date_key not in shifts_by_date:
            shifts_by_date[date_key] = []
        shifts_by_date[date_key].append(shift)
    
    # Create events for each date, combining shift information
    for date_key, date_shifts in shifts_by_date.items():
        event = Event()
        
        # Combine all shift types for the summary
        shift_types = [s['shift_type'] for s in date_shifts]
        day_of_week = date_shifts[0]['day_of_week']  # They all have the same date
        shift_date = date_shifts[0]['date']
        
        # Format the summary to show all shift types
        summary = f"{', '.join(shift_types)} - {day_of_week}"
        event.add('summary', summary)
        
        # All-day events need a DATE value type
        event.add('dtstart', shift_date)
        
        # For all-day events, the end date should be the next day
        # The end date is non-inclusive in the iCalendar spec
        end_date = shift_date + timedelta(days=1)
        event.add('dtend', end_date)
        
        event.add('dtstamp', datetime.now())
        
        # Generate a unique ID for the event
        uid = f"{employee_name.replace(' ', '')}-{shift_date.strftime('%Y%m%d')}@shifts.example.com"
        event.add('uid', uid)
        
        # Add description with details about all employees working that day
        description_parts = [f"Your shifts: {', '.join(shift_types)}"]
        
        # Find all employees working on this date
        coworkers_info = []
        for s in shifts:
            # If it's the same date but not the current employee
            if s['date'] == shift_date and s['employee'].lower() != employee_name.lower():
                coworkers_info.append(f"{s['employee']}: {s['shift_type']}")
        
        # Add coworkers section if any exist
        if coworkers_info:
            description_parts.append("\nCoworkers on this day:")
            for info in sorted(coworkers_info):
                description_parts.append(f"- {info}")
        else:
            description_parts.append("\nNo other employees scheduled on this day.")
        
        # Add Cath Lab on-call information if available
        if cath_lab_shifts:
            cath_lab_employee = None
            for shift in cath_lab_shifts:
                if shift['date'] == shift_date and shift['employee'].lower() != employee_name.lower():
                    cath_lab_employee = shift['employee']
                    break
            
            if cath_lab_employee:
                description_parts.append(f"\nCath Lab On-Call: {cath_lab_employee}")
        
        # Add Electrophysiology on-call information if available
        if ep_shifts:
            ep_employee = None
            for shift in ep_shifts:
                if shift['date'] == shift_date and shift['employee'].lower() != employee_name.lower():
                    ep_employee = shift['employee']
                    break
            
            if ep_employee:
                description_parts.append(f"\nElectrophysiology On-Call: {ep_employee}")
        
        event.add('description', "\n".join(description_parts))
        
        cal.add_component(event)
    
    # Write to file
    with open(output_file, 'wb') as f:
        f.write(cal.to_ical())
    
    return output_file


def save_calendar_file(employee_name):
    """Open a file dialog to choose where to save the calendar file."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    default_filename = f"{employee_name.replace(' ', '_')}_shifts.ics"
    file_path = filedialog.asksaveasfilename(
        title="Save Calendar File",
        defaultextension=".ics",
        filetypes=[("Calendar Files", "*.ics"), ("All Files", "*.*")],
        initialfile=default_filename
    )
    return file_path
def extract_month_year_from_filename(filename):
    """Attempt to extract month and year from the filename."""
    # Example: "ΕΦΗΜΕΡΙΕΣ ΜΑΡΤΙΟΣ 2025.docx"
    month_dict = {
        "ΙΑΝΟΥΑΡΙΟΣ": 1, "ΦΕΒΡΟΥΑΡΙΟΣ": 2, "ΜΑΡΤΙΟΣ": 3, "ΑΠΡΙΛΙΟΣ": 4,
        "ΜΑΙΟΣ": 5, "ΙΟΥΝΙΟΣ": 6, "ΙΟΥΛΙΟΣ": 7, "ΑΥΓΟΥΣΤΟΣ": 8,
        "ΣΕΠΤΕΜΒΡΙΟΣ": 9, "ΟΚΤΩΒΡΙΟΣ": 10, "ΝΟΕΜΒΡΙΟΣ": 11, "ΔΕΚΕΜΒΡΙΟΣ": 12
    }
    
    # Also look for month name in the document content
    month_from_content = None
    if "ΜΑΡΤΙΟΣ" in filename:
        month_from_content = 3
    
    # Default to current month and year if extraction fails
    default_month = datetime.now().month
    default_year = datetime.now().year
    
    try:
        # Try to extract month name and year
        for month_name, month_num in month_dict.items():
            if month_name in filename:
                # Found month, now look for year
                year_match = re.search(r'20\d\d', filename)
                if year_match:
                    year = int(year_match.group())
                    return month_num, year
                return month_num, default_year
        
        # If we found month in content, use that
        if month_from_content:
            year_match = re.search(r'20\d\d', filename)
            if year_match:
                year = int(year_match.group())
                return month_from_content, year
            return month_from_content, default_year
            
    except:
        pass
    
    return default_month, default_year

def main():
    print("Enhanced Employee Shift Calendar Generator")
    print("=========================================")
    
    # Get input file using file browser
    print("Please select the DOCX or DOC file containing shift schedules...")
    input_file = browse_file("Select Main Shift Schedule Document")
    
    if not input_file:
        print("No file selected. Exiting.")
        return
    
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return
    
    print(f"Selected file: {input_file}")
    
    # Extract month and year from filename if possible
    month, year = extract_month_year_from_filename(os.path.basename(input_file))
    
    # Allow user to override detected month/year
    print(f"Detected: Month {month}, Year {year}")
    month_input = input(f"Enter month number (1-12) [default: {month}]: ").strip()
    if month_input and month_input.isdigit() and 1 <= int(month_input) <= 12:
        month = int(month_input)
    
    year_input = input(f"Enter year [default: {year}]: ").strip()
    if year_input and year_input.isdigit() and len(year_input) == 4:
        year = int(year_input)
    
    # Ask if user wants to include Cath Lab on-call shifts
    cath_lab_shifts = None
    include_cath_lab = input("Include Cath Lab on-call shifts? (y/n): ").strip().lower()
    if include_cath_lab == 'y':
        print("Please select the DOCX or DOC file containing Cath Lab on-call schedules...")
        cath_lab_file = browse_file("Select Cath Lab On-Call Schedule")
        if cath_lab_file and os.path.exists(cath_lab_file):
            print(f"Selected Cath Lab file: {cath_lab_file}")
            cath_lab_tables = read_docx_tables(cath_lab_file)
            if cath_lab_tables:
                cath_lab_shifts = []
                for table in cath_lab_tables:
                    cath_shifts = parse_specialty_on_call_table(table)
                    for shift in cath_shifts:
                        shift['shift_type'] = "Cath Lab On-Call"
                    cath_lab_shifts.extend(cath_shifts)
                print(f"Found {len(cath_lab_shifts)} Cath Lab on-call shifts")
            else:
                print("No tables found in the Cath Lab schedule document.")
        else:
            print("No Cath Lab file selected or file not found.")
    
    # Ask if user wants to include Electrophysiology on-call shifts
    ep_shifts = None
    include_ep = input("Include Electrophysiology on-call shifts? (y/n): ").strip().lower()
    if include_ep == 'y':
        print("Please select the DOCX or DOC file containing Electrophysiology on-call schedules...")
        ep_file = browse_file("Select Electrophysiology On-Call Schedule")
        if ep_file and os.path.exists(ep_file):
            print(f"Selected Electrophysiology file: {ep_file}")
            ep_tables = read_docx_tables(ep_file)
            if ep_tables:
                ep_shifts = []
                for table in ep_tables:
                    electro_shifts = parse_specialty_on_call_table(table)
                    for shift in electro_shifts:
                        shift['shift_type'] = "Electrophysiology On-Call"
                    ep_shifts.extend(electro_shifts)
                print(f"Found {len(ep_shifts)} Electrophysiology on-call shifts")
            else:
                print("No tables found in the Electrophysiology schedule document.")
        else:
            print("No Electrophysiology file selected or file not found.")
    
    # Read and parse the main document
    print(f"Reading main file: {input_file}")
    tables = read_docx_tables(input_file)
    
    if not tables:
        print("No tables found in the document.")
        return
    
    # Parse shifts from both tables
    all_shifts = []
    
    # Process first table (if exists)
    if len(tables) >= 1:
        print("Parsing first table (Regular/On-Call shifts)...")
        first_table_shifts = parse_first_table(tables[0], month, year)
        all_shifts.extend(first_table_shifts)
        print(f"Found {len(first_table_shifts)} shifts in first table")
    
    # Process second table (if exists)
    if len(tables) >= 2:
        print("Parsing second table (Μεγάλη/Μικρή/ΤΕΠ shifts)...")
        second_table_shifts = parse_second_table(tables[1], month, year)
        all_shifts.extend(second_table_shifts)
        print(f"Found {len(second_table_shifts)} shifts in second table")
    
    if not all_shifts:
        print("No shifts found in any table!")
        return
    
    # Get unique employee names across all shifts
    all_employees = sorted(set(shift['employee'] for shift in all_shifts))
    print(f"Found {len(all_shifts)} total shift assignments for {len(all_employees)} employees:")
    
    # Display employee list
    for i, emp in enumerate(all_employees, 1):
        print(f"{i}. {emp}")
    
    # Ask user which employee to generate calendar for
    while True:
        employee_choice = input("\nEnter employee name or number (or 'all' for all employees): ").strip()
        
        if employee_choice.lower() == 'all':
            # Generate calendars for all employees
            for employee in all_employees:
                output_file = save_calendar_file(employee)
                if output_file:
                    result = create_calendar_for_employee(all_shifts, employee, output_file, 
                                                          cath_lab_shifts, ep_shifts)
                    if result:
                        print(f"Calendar for {employee} created successfully: {output_file}")
            break
        
        # Check if user entered a number
        elif employee_choice.isdigit() and 1 <= int(employee_choice) <= len(all_employees):
            employee_name = all_employees[int(employee_choice) - 1]
        else:
            # Assume user entered a name
            employee_name = employee_choice
            # Check if name exists in our list
            if employee_name.lower() not in [emp.lower() for emp in all_employees]:
                closest_match = None
                for emp in all_employees:
                    if employee_name.lower() in emp.lower():
                        closest_match = emp
                        break
                
                if closest_match:
                    confirm = input(f"Did you mean '{closest_match}'? (y/n): ").lower()
                    if confirm == 'y':
                        employee_name = closest_match
                    else:
                        print("Please try again.")
                        continue
                else:
                    print(f"Employee '{employee_name}' not found. Please try again.")
                    continue
        
        # Get output file location using file browser
        print(f"Select where to save the calendar file for {employee_name}...")
        output_file = save_calendar_file(employee_name)
        
        if not output_file:
            print("Calendar save operation cancelled.")
            retry = input("Would you like to select a different employee? (y/n): ").lower()
            if retry == 'y':
                continue
            else:
                break
        
        # Create calendar with additional shifts
        result = create_calendar_for_employee(all_shifts, employee_name, output_file, 
                                              cath_lab_shifts, ep_shifts)
        if result:
            print(f"Calendar created successfully: {output_file}")
            
        # Ask if user wants to create another calendar
        another = input("Would you like to create a calendar for another employee? (y/n): ").lower()
        if another != 'y':
            break

if __name__ == "__main__":
    main()

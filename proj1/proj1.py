import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook

course_dict = {}
def generate_attendance_sheets(allocation_file, student_data_file, output_file="Attendance_Sheets.xlsx"):
    """
    Generates attendance sheets based on the allocation data and student information.
    """
    # Load the exam allocation file
    allocation_df = pd.read_excel(allocation_file)

    # Load the student data and create a dictionary mapping roll numbers to names
    students_df = student_data_file
    student_dict = pd.Series(students_df.Name.values, index=students_df.Roll).to_dict()

    # Initialize a workbook for attendance sheets
    workbook = Workbook()
    workbook.remove(workbook.active)  # Remove the default sheet created by Workbook()

    # Function to add a new sheet with attendance data
    def add_attendance_sheet(workbook, sheet_name, roll_numbers):
        sheet = workbook.create_sheet(title=sheet_name[:31])  # Limit sheet name to 31 characters
        sheet['A1'] = 'Roll_No'
        sheet['B1'] = 'Name'
        sheet['C1'] = 'Signature'

        for idx, roll in enumerate(roll_numbers, start=2):  # Start from row 2
            roll = roll.strip()
            sheet[f'A{idx}'] = roll
            sheet[f'B{idx}'] = student_dict.get(roll, "Unknown")
            sheet[f'C{idx}'] = ''  # Signature column

    # Process each row in the allocation DataFrame
    for _, row in allocation_df.iterrows():
        date = pd.to_datetime(row['Date']).strftime("%d_%m_%Y")
        course_code = row['course_code']
        room_no = row['Room']
        time_slot = row['Time'].lower()  # Ensure lowercase for consistency
        roll_numbers = row['Roll_list'].split(';')

        # Generate a unique sheet name
        sheet_name = f"{date}{course_code}{room_no}_{time_slot}"
        add_attendance_sheet(workbook, sheet_name, roll_numbers)

    # Save the workbook
    workbook.save(output_file)
    print(f"Attendance sheets saved to '{output_file}'")
# Define a function to read data from an Excel file and process it based on sheet name
def read_excel_data(file_path, sheet_name, header=1):
    """
    Reads data from the given Excel file and sheet, returning a DataFrame.
    """
    return pd.read_excel(file_path, sheet_name=sheet_name, header=header)

# Function to extract and organize student data by course code
def organize_students_by_course(df):
    """
    Organizes student roll numbers into a dictionary by course code.
    """
    
    for _, row in df.iterrows():
        course_code = row['course_code']
        rollno = row['rollno']
        if course_code in course_dict:
            course_dict[course_code].append(rollno)
        else:
            course_dict[course_code] = [rollno]
    
    for course_code in course_dict:
        course_dict[course_code].sort()  # Sort roll numbers for each course
    course_dict2=course_dict
    return course_dict2
 
# Function to separate rooms by type (floor vs LT rooms) and sort them
def process_rooms(df):
    """
    Separates the rooms into floor-based and LT rooms, sorting them by floor and capacity.
    """
    df.head()
    floor_rooms = df[df['Room No.'].astype(str).str.match(r'^\d+')]
    lt_rooms = df[df['Room No.'].astype(str).str.startswith('LT')]

    floor_rooms['Floor'] = floor_rooms['Room No.'].astype(str).str[0]
    sorted_floor_rooms = floor_rooms.sort_values(by=['Floor', 'Exam Capacity'], ascending=[True, False])
    
    lt_rooms['LT_Floor'] = lt_rooms['Room No.'].astype(str).str[2]
    sorted_lt_rooms = lt_rooms.sort_values(by=['LT_Floor', 'Room No.'], ascending=[True, True])

    # Concatenate sorted floor and LT rooms
    sorted_rooms = pd.concat([sorted_floor_rooms, sorted_lt_rooms]).drop(columns=['Floor', 'LT_Floor'])
    print(sorted_rooms)
    return sorted_rooms

# Function to process and update room capacities
def update_room_capacities(df, buffer):
    """
    Updates room capacities after applying the buffer.
    """
    df['Remaining Capacity'] = df['Exam Capacity'] - buffer
    return df

# Function to parse and format the exam timetable
def process_exam_timetable(df):
    """
    Converts the exam timetable into a structured dictionary.
    """
    exam_timetable = {}
    for _, row in df.iterrows():
        date = row['Date']
        morning_courses = row['Morning']
        evening_courses = row['Evening']

        exam_timetable[f"{date}_morning"] = morning_courses
        exam_timetable[f"{date}_evening"] = evening_courses

    # Format courses in the timetable
    print(exam_timetable)
    formatted_timetable = {}
    for key, courses in exam_timetable.items():
        formatted_timetable[key] = [courses]
        
    exam_timetable =formatted_timetable
    for key, courses in exam_timetable.items():
       if courses[0] != 'NO EXAM':  # Skip if no exam
        # Split courses, sort them using lambda based on course_dict length, and rejoin
         sorted_courses = sorted(courses[0].split('; '), key=lambda course: len(course_dict.get(course, [])), reverse=True)
         exam_timetable[key] = ['; '.join(sorted_courses)]
    print(exam_timetable)
    return exam_timetable

# Function to get the weekday name from a date string
def get_weekday_from_date(date_str):
    """
    Extracts the weekday name from a date string.
    """
    try:
        date_obj = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
    except ValueError:
        raise ValueError(f"Unexpected date format: {date_str}")
    return date_obj.strftime('%A')

# Function to allocate students to rooms for the 'dense' scenario
def allocate_students_dense(exam_timetable, students_data, rooms_data, buffer):
    """
    Allocates students to rooms for dense allocation, considering room capacity and students per course.
    """
    data = []
    for exam_key, course_list in exam_timetable.items():
        if 'NO EXAM' in course_list:
            continue
        
        date_part, time_part = exam_key.split('_')
        day_part = get_weekday_from_date(date_part)  # Get weekday

        courses = course_list[0].split('; ')  # Get list of courses
        for course in courses:
            students = students_data.get(course, [])

            # Iterate over rooms and allocate students based on remaining capacity
            for i, room in rooms_data.iterrows():
                room_no = room['Room No.']
                remaining_capacity = room['Remaining Capacity']
                allocated_students = []

                while students and remaining_capacity > 0:
                    # Allocate students to the room until capacity is filled
                    allocated_students.append(students.pop(0))
                    remaining_capacity -= 1

                # Update remaining capacity in the rooms DataFrame
                rooms_data.at[i, 'Remaining Capacity'] = remaining_capacity

                # If students were allocated, store the results
                if allocated_students:
                    data.append({
                        'Date': date_part,
                        'Day': day_part,
                        'Time': time_part,
                        'course_code': course,
                        'Room': room_no,
                        'Allocated_students_count': len(allocated_students),
                        'Roll_list': '; '.join(allocated_students)
                    })
                if not students:
                    break
    return pd.DataFrame(data)
def allocate_students_sparse(exam_timetable, students_data, rooms_data):
    """
    Allocates students to rooms for sparse allocation, balancing students across multiple rooms.
    """
    df2 = rooms_data.copy(deep=True)

    # Initialize a list to hold the data for the DataFrame
    data = []

    def get_floor(room_no):
     room_no_str = str(room_no)  # Explicitly convert room number to a string
     if room_no_str.startswith('LT'):
        return 6  # Assign LT rooms to a higher floor number
     else:
        # Safeguard: Handle cases where the room number might not start with a digit
        return int(room_no_str[0]) if room_no_str[0].isdigit() else 0  # First character as floor number, default to 0 if not a digit

    # Create a new column for the floor number
    df2['Floor'] = df2['Room No.'].apply(get_floor)

    # Sort by Floor (ascending) and then by Remaining Capacity (descending)
    df2_sorted = df2.sort_values(by=['Floor', 'Remaining Capacity'], ascending=[True, False]).reset_index(drop=True)

    # Drop the 'Floor' column if no longer needed
    df2_sorted = df2_sorted.drop(columns=['Floor'])
    df2 = df2_sorted

    # Calculate 'rem cap 1' and 'rem cap 2' columns
    df2['rem cap 1'] = np.ceil(df2['Remaining Capacity'] / 2).astype(int)  # Ceiling of remaining capacity divided by 2
    df2['rem cap 2'] = df2['Remaining Capacity'] - df2['rem cap 1']  # Remaining capacity minus rem cap 1
    dfx = df2.copy(deep=True)

    # Start allocating students to rooms
    for exam_key, course_list in exam_timetable.items():
        if 'NO EXAM' in course_list:
            continue  # Skip if no exam for that slot

        dfx = df2.copy(deep=True)
        i = 0  # Pointer for rem cap 1
        j = 0  # Pointer for rem cap 2

        # Split the exam_key to get date and time
        date_part, time_part = exam_key.split('_')
        day_part = get_day_from_date(date_part)  # Get the weekday

        courses = course_list[0].split('; ')  # Get the courses for that exam session

        for course in courses:
            students = students_data.get(course, [])  # Get the list of students for the course

            # Determine which pointer to use based on their values
            current_pointer = i if i <= j else j  # Choose i if i <= j, otherwise choose j
            current_rem_cap = 'rem cap 1' if current_pointer == i else 'rem cap 2'

            while students:
                allocated_students = []

                # Allocate students using the selected pointer
                if dfx.at[current_pointer, current_rem_cap] > 0:
                    # Allocate to the selected room
                    while students and dfx.at[current_pointer, current_rem_cap] > 0:
                        allocated_students.append(students.pop(0))
                        dfx.at[current_pointer, current_rem_cap] -= 1  # Decrease remaining capacity

                    # If we allocated students, store the result
                    if allocated_students:
                        data.append({
                            'Date': date_part,
                            'Day': day_part,
                            'Time': time_part,
                            'course_code': course,
                            'Room': dfx.at[current_pointer, 'Room No.'],
                            'Allocated_students_count': len(allocated_students),
                            'Roll_list': '; '.join(allocated_students)
                        })

                # After attempting to allocate, check if we need to increment the pointer
                if current_pointer == i and dfx.at[i, 'rem cap 1'] == 0:
                    i += 1  # Move to the next room for rem cap 1
                    current_pointer = i

                elif current_pointer == j and dfx.at[j, 'rem cap 2'] == 0:
                    j += 1  # Move to the next room for rem cap 2
                    current_pointer = j

    # Create a DataFrame from the collected data
    return pd.DataFrame(data)
# Main function to handle the logic flow based on the user input
def main(file_path, buffer, density_type):
    # Read the necessary data
    ip_1 = pd.read_excel(file_path, sheet_name='ip_1', header=1)
    ip_3 = pd.read_excel(file_path, sheet_name='ip_3', header=0)
    ip_2 = pd.read_excel(file_path, sheet_name='ip_2',header=1)
    ip_4 = pd.read_excel(file_path, sheet_name='ip_4', header=0)
    students_data = organize_students_by_course(read_excel_data(file_path, 'ip_1'))
    print(students_data)
    rooms_data = process_rooms(ip_3)
    rooms_data = update_room_capacities(rooms_data, buffer)
    exam_timetable = process_exam_timetable(ip_2)

    # Allocate students based on the density type (dense or sparse)
    if density_type == 'dense':
        result_df = allocate_students_dense(exam_timetable, students_data, rooms_data, buffer)
    else:
        result_df = allocate_students_sparse(exam_timetable, students_data, rooms_data)

    # Save the final result to an Excel file
    excel_file_name = 'exam_allocation.xlsx'
    result_df.to_excel(excel_file_name, index=False)
    print(f"Excel file '{excel_file_name}' created successfully.")
    generate_attendance_sheets(excel_file_name, ip_4)
# User interaction for buffer value and density type
def user_input():
    buffer = int(input("Enter a buffer value (0-5): "))
    while buffer < 0 or buffer > 5:
        buffer = int(input("Invalid input. Enter a buffer value (0-5): "))

    density_input = input("Enter '1' for dense or '2' for sparse: ")
    while density_input not in ['1', '2']:
        density_input = input("Invalid input. Please enter '1' for dense or '2' for sparse: ")

    density_type = 'dense' if density_input == '1' else 'sparse'
    print(f"You selected: {density_type}")
    return buffer, density_type

# Function to generate attendance sheets from the allocated exam data

    
# Final call to the user input function and processing
if _name_ == '_main_':
    buffer, density_type = user_input()
    file_path = 'sir_input.xlsx'  # Hardcoded file path
    main(file_path, buffer, density_type)
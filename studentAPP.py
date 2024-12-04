import openpyxl

def create_excel_file(file_name):
    # Create a new workbook
    workbook = openpyxl.Workbook()

    # Get the active sheet (by default, there's one sheet named 'Sheet')
    sheet = workbook.active

    # Create the header row
    header = ["Name", "ID", "Attendance", "Email", "Phone", "Address", "Course", "Year"]
    sheet.append(header)

    # Save the workbook
    workbook.save(file_name)

def add_student_to_excel(file_name):
    # Load the existing workbook
    workbook = openpyxl.load_workbook(file_name)

    # Get the active sheet
    sheet = workbook.active

    # Get student data from the user
    name = input("Enter the student's name: ")
    student_id = int(input("Enter the student's ID: "))
    attendance = int(input("Enter the student's attendance percentage: "))
    email = input("Enter the student's email: ")
    phone = input("Enter the student's phone number: ")
    address = input("Enter the student's address: ")
    course = input("Enter the student's course: ")
    year = input("Enter the student's year: ")

    # Append the student data to the sheet
    student_data = [name, student_id, attendance, email, phone, address, course, year]
    sheet.append(student_data)

    # Save the workbook to update the Excel file
    workbook.save(file_name)

def read_student_data(file_name):
    # Load the existing workbook
    workbook = openpyxl.load_workbook(file_name)

    # Get the active sheet
    sheet = workbook.active

    # Iterate through rows to read student data
    for row in sheet.iter_rows(min_row=2, values_only=True):
        name, student_id, attendance, email, phone, address, course, year = row
        print(f"Name: {name}, ID: {student_id}, Attendance: {attendance}%")
        print(f"Email: {email}, Phone: {phone}, Address: {address}")
        print(f"Course: {course}, Year: {year}")
        print()

if __name__ == "__main__":
    file_name = "C:\\Users\\ASUS\\PycharmProjects\\pythonProject1\\StudentDatabaseFINAL.xlsx"  # Change this to your desired file name

    create_excel_file(file_name)

    while True:
        print("Options:")
        print("1. Add a student record")
        print("2. Read student data")
        print("3. Exit")
        choice = input("Enter your choice (1/2/3): ")

        if choice == "1":
            add_student_to_excel(file_name)
        elif choice == "2":
            read_student_data(file_name)
        elif choice == "3":
            print("Exiting the program.")
            break
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")

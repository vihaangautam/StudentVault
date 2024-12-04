StudentVault

The StudentVault is a student database management system based in Python. It makes the process of storing and retrieval of student information easy. The system uses excel files for data storage, which keeps it lightweight and very simple to use.

Features-
Add Student Record: Easily enter student details such as name, ID, attendance, contact details, and academic information.
Retrieve Student Records: View all stored data of students in an organized manner.
Excel Integrator: Store their students' information in an Excel file for easy access and portability.

Requirements-
Python 3.6+
openpyxl library
Install the required library from pip:
pipe install openpyxl

Usage-
1. Create an Excel file
On the first execution, the system will automatically create an Excel file called StudentDatabaseFINAL.xlsx at the given directory. The file shall be used to contain all records of students.
2. Run the Program
Run the script using Python:
python student_database_code.py
3. Select an Option
The program offers three options:

Add student record:
Enter the name, ID, attendance, email, phone, address, course, and year of the student.
The data will get appended to the excel file.

Read student data:
All student records as saved in the excel file will be displayed in readable format.

Exit:
Will exit from the program.

Example Interaction-

Options:
1. Add student record
2. Read student data
3. Exit
Enter your choice (1/2/3): 1
Enter the student's name: John Doe
Enter the student's ID: 101
Enter the student's attendance percentage: 95
Enter the student's email: johndoe@example.com
Enter the student's phone number: 1234567890
Enter the student's address: 123 Main St
Enter the student's course: Computer Science
Enter the student's year: 3
Record added successfully.

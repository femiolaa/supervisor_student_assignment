import pandas as pd

# Load Excel files
students_file = "SIWES 2025.xlsx"
supervisors_file = "SIWES SUPERVISORS.xlsx"

# Read the data
students_df = pd.read_excel(students_file)
supervisors_df = pd.read_excel(supervisors_file)

# Assuming the first column contains student names and the first column of supervisors
students = students_df.iloc[:, 0].tolist()
supervisors = supervisors_df.iloc[:, 0].tolist()

# Calculate the number of supervisors
num_supervisors = len(supervisors)

# Assign supervisors to students in a round-robin manner
assignments = {}
for i, student in enumerate(students):
    supervisor = supervisors[i % num_supervisors]  # Assign by index
    assignments[student] = supervisor

# Create a DataFrame to save results
assignments_df = pd.DataFrame(assignments.items(), columns=['Student', 'Supervisor'])

# Save to a new Excel file
assignments_df.to_excel("Student_Supervisor_Assignment.xlsx", index=False)
print("Assignments saved to 'Student_Supervisor_Assignment.xlsx'")

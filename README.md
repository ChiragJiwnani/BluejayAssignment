#Employee Timecard Analysis
This Python script analyzes an Excel file containing employee timecard data. The analysis includes identifying employees who meet the following criteria:

7 Consecutive Unique Days:

Identifies employees who have worked for 7 consecutive unique days.
Less Than 10 Hours Between Shifts (Greater Than 1 Hour):

Identifies employees who have less than 10 hours between shifts but greater than 1 hour.
More Than 14 Hours in a Single Shift:

Identifies employees who have worked for more than 14 hours in a single shift.

#Output
The script generates the following outputs in the 'output.txt' file:

Employees with 7 consecutive unique days.
Employees with less than 10 hours between shifts but greater than 1 hour.
Employees with more than 14 hours in a single shift.
Additionally, the output file includes the count of occurrences of short shifts and long shifts for each employee.

#Notes
The script assumes that the input Excel file has the necessary columns: 'Time', 'Timecard Hours (as Time)', 'Employee Name', 'Position ID'.
Ensure that the time-related columns are in the correct format.

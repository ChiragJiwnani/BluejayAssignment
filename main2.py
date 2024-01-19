import pandas as pd
from datetime import datetime, timedelta

def analyze_excel_file(file_path, output_file='output.txt'):
    # Read the Excel file into a pandas DataFrame
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return

    # Assuming 'Time' column is in datetime format, if not, convert it to datetime
    if not pd.api.types.is_datetime64_any_dtype(df['Time']):
        df['Time'] = pd.to_datetime(df['Time'])

    # Convert 'Timecard Hours (as Time)' to time format
    df['hours_worked'] = pd.to_datetime(df['Timecard Hours (as Time)'], format='%H:%M').dt.hour + pd.to_datetime(df['Timecard Hours (as Time)'], format='%H:%M').dt.minute / 60

    # Sort the DataFrame by 'Employee Name' and 'Time' for easier analysis
    df = df.sort_values(by=['Employee Name', 'Time'])

    # Open the output file in write mode
    with open(output_file, 'w') as output:
        consecutive_days_list = []
        short_shifts_list = []
        long_shifts_list = []

        # Count occurrences of short shifts
        short_shifts_count = df[(df['hours_worked'] > 1) & (df['hours_worked'] < 10)].groupby('Employee Name').size().reset_index(name='Short Shifts Count')

        # Count occurrences of long shifts
        long_shifts_count = df[df['hours_worked'] > 14].groupby('Employee Name').size().reset_index(name='Long Shifts Count')

        # Iterate through the DataFrame to analyze the conditions
        for name, group in df.groupby('Employee Name'):
            consecutive_days_set = set()
            short_shifts_set = set()
            long_shifts_set = set()

            for index, row in group.iterrows():
                time = row['Time']
                hours_worked = row['hours_worked']

                # Check for consecutive unique days
                if consecutive_days_set and (time - max(consecutive_days_set)).days == 1:
                    consecutive_days_set.add(time)
                else:
                    consecutive_days_set = {time}

                if len(consecutive_days_set) == 7:
                    consecutive_days_list.append(f"{name} ({row['Position ID']}): 7 consecutive unique days.")
                    # Clear the set after identifying 7 consecutive unique days
                    consecutive_days_set.clear()

                # Check for less than 10 hours between shifts but greater than 1 hour
                if short_shifts_set and (time - max(short_shifts_set)).total_seconds() / 3600 < 10:
                    short_shifts_set.add(time)
                else:
                    short_shifts_set = {time}

                if hours_worked > 1 and len(short_shifts_set) == 2:
                    short_shifts_list.append(f"{name} ({row['Position ID']}): Less than 10 hours between shifts but greater than 1 hour.")
                    # Clear the set after identifying the condition
                    short_shifts_set.clear()

                # Check if hours worked are more than 14 in a single shift
                if long_shifts_set and (time - max(long_shifts_set)).total_seconds() / 3600 > 14:
                    long_shifts_set.add(time)
                else:
                    long_shifts_set = {time}

                if hours_worked > 14 and len(long_shifts_set) == 1:
                    long_shifts_list.append(f"{name} ({row['Position ID']}): More than 14 hours in a single shift.")
                    # Clear the set after identifying the condition
                    long_shifts_set.clear()

        # Write divisions to the output file
        output.write("Employees with 7 consecutive unique days:\n")
        output.write("\n".join(consecutive_days_list) + "\n\n")

        output.write("Employees with less than 10 hours between shifts but greater than 1 hour:\n")
        output.write("\n".join(short_shifts_list) + "\n\n")

        output.write("Employees with more than 14 hours in a single shift:\n")
        output.write("\n".join(long_shifts_list) + "\n")

        # Write the count of occurrences of short shifts for each employee
        output.write("Count of occurrences of short shifts for each employee:\n")
        output.write(short_shifts_count.to_string(index=False) + "\n")

        # Write the count of occurrences of long shifts for each employee
        output.write("Count of occurrences of long shifts for each employee:\n")
        output.write(long_shifts_count.to_string(index=False) + "\n")

# Assuming the input Excel file is named 'Assignment_Timecard.xlsx'
analyze_excel_file('Assignment_Timecard.xlsx', output_file='output.txt')

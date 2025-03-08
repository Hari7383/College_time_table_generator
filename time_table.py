import pandas as pd
import random

def load_staff_data(file_path):
    return pd.read_excel(file_path)

def generate_timetable(staff_data):
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    hours_per_day = 5
    timetable = {day: [] for day in days}
    
    for day in days:
        available_staff = staff_data.copy()
        for hour in range(hours_per_day):
            if available_staff.empty:
                available_staff = staff_data.copy()
            
            staff = available_staff.sample(n=1)
            staff_name = staff.iloc[0]['Staff Name']
            department = staff.iloc[0]['Department']
            subject = staff.iloc[0]['Handling Subject']
            
            timetable[day].append({
                'Hour': hour + 1,
                'Staff': staff_name,
                'Department': department,
                'Subject': subject
            })
            
            available_staff = available_staff[available_staff['Staff Name'] != staff_name]
            if random.random() < 0.3 and hour < hours_per_day - 1:  # 30% chance of continuous hours
                timetable[day].append({
                    'Hour': hour + 2,
                    'Staff': staff_name,
                    'Department': department,
                    'Subject': subject
                })
                hour += 1
    
    return timetable

def save_timetable_to_excel(timetable, output_file):
    with pd.ExcelWriter(output_file) as writer:
        for day, slots in timetable.items():
            df = pd.DataFrame(slots)
            df.to_excel(writer, sheet_name=day, index=False)

def main():
    input_file = 'staff_data.xlsx'  # Update with your file path
    output_file = 'timetable.xlsx'
    
    staff_data = load_staff_data(input_file)
    timetable = generate_timetable(staff_data)
    save_timetable_to_excel(timetable, output_file)
    print(f"Timetable saved to {output_file}")

if __name__ == "__main__":
    main()
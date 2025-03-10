import pandas as pd
import random

def load_staff_data(file_path):
    df = pd.read_excel(file_path, sheet_name="Sheet1")
    df = df[['STAFF NAMES', 'CLASS', 'SUBJECTS', 'HOURS']]
    df = df.dropna(subset=['STAFF NAMES'])
    return df

def generate_timetable(staff_data):
    days = [f"Day {i+1}" for i in range(20)]  # 4 weeks (5 days per week)
    hours_per_day = 5
    timetable = {day: [] for day in days}
    
    for day in days:
        available_staff = staff_data.copy()
        for hour in range(hours_per_day):
            if available_staff.empty:
                available_staff = staff_data.copy()
            
            staff = available_staff.sample(n=1)
            staff_name = staff.iloc[0]['STAFF NAMES']
            class_name = staff.iloc[0]['CLASS']
            subject = staff.iloc[0]['SUBJECTS']
            
            timetable[day].append({
                'Hour': hour + 1,
                'Staff': staff_name,
                'Class': class_name,
                'Subject': subject
            })
            
            available_staff = available_staff[available_staff['STAFF NAMES'] != staff_name]
            if random.random() < 0.3 and hour < hours_per_day - 1:  # 30% chance of continuous hours
                timetable[day].append({
                    'Hour': hour + 2,
                    'Staff': staff_name,
                    'Class': class_name,
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
    input_file = 'staffTT.xlsx'  # Update with your file path
    output_file = 'timetable.xlsx'
    
    staff_data = load_staff_data(input_file)
    timetable = generate_timetable(staff_data)
    save_timetable_to_excel(timetable, output_file)
    print(f"Timetable saved to {output_file}")

if __name__ == "__main__":
    main()

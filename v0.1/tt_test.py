import pandas as pd
import random

def load_staff_data(file_path):
    df = pd.read_excel(file_path, sheet_name="Sheet1")
    df = df[['STAFF NAMES', 'CLASS', 'SUBJECTS']]
    df = df.dropna(subset=['STAFF NAMES'])
    return df

def generate_timetable(staff_data):
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    hours_per_day = 5
    staff_list = staff_data['STAFF NAMES'].unique()
    
    timetable = {staff: {day: [None] * hours_per_day for day in days} for staff in staff_list}
    
    for staff_name in staff_list:
        staff_subjects = staff_data[staff_data['STAFF NAMES'] == staff_name]
        
        for day in days:
            teaching_hours = random.sample(range(hours_per_day), 4)  # 4 teaching hours, 1 free hour
            selected_subjects = staff_subjects.sample(n=4, replace=True)  # Pick multiple subjects and classes
            subjects_classes = selected_subjects[['CLASS', 'SUBJECTS']].to_dict(orient='records')
            
            sub_index = 0
            for hour in range(hours_per_day):
                if hour in teaching_hours:
                    chosen_entry = subjects_classes[sub_index]
                    timetable[staff_name][day][hour] = f"{chosen_entry['CLASS']} - {chosen_entry['SUBJECTS']}"
                    sub_index += 1
                else:
                    timetable[staff_name][day][hour] = "Free Hour"
    
    return timetable

def save_timetable_to_excel(timetable, output_file):
    with pd.ExcelWriter(output_file) as writer:
        for staff, schedule in timetable.items():
            df = pd.DataFrame(schedule).T.reset_index()
            df.columns = ["Day", "Hour 1", "Hour 2", "Hour 3", "Hour 4", "Hour 5"]
            df.loc[-1] = [staff] + ["" for _ in range(5)]  # Staff name as header row
            df.index = df.index + 1  # Shift index
            df = df.sort_index()
            df.to_excel(writer, sheet_name=staff[:30], index=False)

def main():
    input_file = 'tt/staffTT.xlsx'
    output_file = 'tt/timetable.xlsx'
    
    staff_data = load_staff_data(input_file)
    timetable = generate_timetable(staff_data)
    save_timetable_to_excel(timetable, output_file)
    print(f"Timetable saved to {output_file}")

if __name__ == "__main__":
    main()

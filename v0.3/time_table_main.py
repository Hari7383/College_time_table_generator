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
    
    # Dictionary to store the timetable ensuring no class collision
    timetable = {staff: {day: [None] * hours_per_day for day in days} for staff in staff_list}
    
    # Track which classes are already scheduled at a given hour
    scheduled_classes = {day: {hour: set() for hour in range(hours_per_day)} for day in days}

    for staff_name in staff_list:
        staff_subjects = staff_data[staff_data['STAFF NAMES'] == staff_name]

        for day in days:
            teaching_hours = random.sample(range(hours_per_day), 4)  # 4 teaching hours, 1 free hour
            available_subjects = staff_subjects.sample(frac=1).to_dict(orient='records')  # Shuffle subjects
            
            sub_index = 0
            for hour in range(hours_per_day):
                if hour in teaching_hours and sub_index < len(available_subjects):
                    chosen_entry = available_subjects[sub_index]
                    class_name = chosen_entry['CLASS']
                    subject_class_pair = f"{class_name} - {chosen_entry['SUBJECTS']}"
                    
                    # Ensure no class collision at this hour
                    if class_name not in scheduled_classes[day][hour]:  # Check if class is already occupied
                        timetable[staff_name][day][hour] = subject_class_pair
                        scheduled_classes[day][hour].add(class_name)  # Mark class as used
                        sub_index += 1
                    else:
                        timetable[staff_name][day][hour] = "Free Hour"  # If collision, give free hour
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
    input_file = 'tt/staffTT_edited.xlsx'  # Replace with your file path
    output_file = 'tt/output_timetable.xlsx'  # Replace Output timetable file path
    
    staff_data = load_staff_data(input_file)
    timetable = generate_timetable(staff_data)
    save_timetable_to_excel(timetable, output_file)
    print(f"Timetable saved to {output_file}")

if __name__ == "__main__":
    main()
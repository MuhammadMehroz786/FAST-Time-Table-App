import streamlit as st
import pandas as pd

# Load the Excel file
@st.cache_resource
def load_excel_file():
    file_path = '/Users/apple/Desktop/Fast 3rd Sem/Timetable Streamlit/timetable.xlsx'
    try:
        excel_data = pd.ExcelFile(file_path)
    except FileNotFoundError:
        st.error("The timetable file was not found.")
        return None
    return excel_data

# Load the data for a specific sheet
@st.cache_data
def load_sheet(day):
    excel_data = load_excel_file()
    if excel_data is None:
        return None
    try:
        df = pd.read_excel(excel_data, sheet_name=day)
        
        # Extract slots and information
        slots = df.iloc[1, 1:11].tolist()  # Slots from second row, second to K column
        venues = df.iloc[4:58, 0].tolist()  # Venue names from 5A to 57A
        info = df.iloc[4:58, 1:11]  # Information from 5B to 5K and 57B to 57K
        
        # List of electives
        electives = {'FOM', 'FOA', 'POE'}

        # Process the information to extract Subject, Department, and Class
        timetable_data = []
        for index, venue in enumerate(venues):
            for slot in slots:
                cell_value = info.iloc[index][slots.index(slot)]
                if pd.notna(cell_value):
                    parts = cell_value.splitlines()  # Split by new line
                    if len(parts) >= 2:
                        subject_department_class = parts[0].strip()  # First line contains Subject, Department, and Class
                        instructor = parts[1].strip()  # Second line contains Instructor

                        # Extract the subject
                        subject_department_class_parts = subject_department_class.split()
                        subject = subject_department_class_parts[0]  # Assume the subject is the first word

                        # Handle elective and lab entries separately
                        if subject in electives:
                            # Handle elective entries
                            subject_department_class_parts = subject_department_class.split()
                            if len(subject_department_class_parts) >= 2:
                                subject = ' '.join(subject_department_class_parts[:-1])  # All except last part are the subject
                                department_class = subject_department_class_parts[-1]  # Last part is Department-Class
                                if '-' in department_class:
                                    department, class_name = department_class.split('-')  # Split into Department and Class
                                else:
                                    department = department_class
                                    class_name = ''  # Fallback if no class is present
                                
                                timetable_data.append({
                                    'Time': slot,
                                    'Venue': venue,
                                    'Subject': subject,
                                    'Department': department,
                                    'Class': class_name.strip()  # Ensure no extra spaces
                                })
                        elif "Lab" in subject_department_class:
                            # Handle lab entries specifically
                            subject_department_class_parts = subject_department_class.split()
                            if len(subject_department_class_parts) >= 2:
                                subject = ' '.join(subject_department_class_parts[:-1])  # All except last part are the subject
                                department_class = subject_department_class_parts[-1]  # Last part is Department-Class
                                if '-' in department_class:
                                    department, class_name = department_class.split('-')  # Split into Department and Class
                                else:
                                    department = department_class
                                    class_name = ''  # Fallback if no class is present
                                
                                timetable_data.append({
                                    'Time': slot,
                                    'Venue': venue,
                                    'Subject': subject,
                                    'Department': department,
                                    'Class': class_name.strip()  # Ensure no extra spaces
                                })
                        else:
                            # Handle regular classes
                            if len(subject_department_class_parts) >= 2:
                                subject = subject_department_class_parts[0]  # First part is Subject
                                department_class = subject_department_class_parts[1]  # Second part is Department-Class
                                if '-' in department_class:
                                    department, class_name = department_class.split('-')  # Split into Department and Class
                                else:
                                    department = department_class
                                    class_name = ''  # Fallback if no class is present
                                
                                timetable_data.append({
                                    'Time': slot,
                                    'Venue': venue,
                                    'Subject': subject,
                                    'Department': department,
                                    'Class': class_name.strip()  # Ensure no extra spaces
                                })
        
        return pd.DataFrame(timetable_data)

    except Exception as e:
        st.error(f"Error loading sheet {day}: {e}")
        return None

def show_schedule():
    excel_data = load_excel_file()  # Load the Excel data
    if excel_data is None:
        return

    day = st.selectbox("Select a Day", excel_data.sheet_names)  # Ask for the day
    
    # Section for regular classes
    department_class_input = st.text_input("Enter Department and Class (e.g., 'BCS 3F'):").strip().upper()  # Single input field for non-elective classes

    if department_class_input:  # Only proceed if input is entered
        # Split the input into department and class
        parts = department_class_input.split()
        if len(parts) >= 2:
            department = parts[0]  # First part is Department
            class_name = parts[1]  # Second part is Class
        else:
            st.write("Please enter both department and class in the format 'Department Class'.")
            return

        schedule_df = load_sheet(day)  # Load the schedule for the specified day
        if schedule_df is not None:
            # Filter and sort the DataFrame for the specified department and class
            filtered_schedule = schedule_df[
                (schedule_df['Department'] == department) & 
                (schedule_df['Class'] == class_name)
            ]
            
            # Ensure 'Time' is sorted correctly. This assumes 'Time' is a string that represents time in a sortable format.
            filtered_schedule = filtered_schedule.sort_values(by='Time')
            
            if not filtered_schedule.empty:
                st.write(f"### Timetable for Department **{department}** and Class **{class_name}** on **{day}**:")
                
                for index, row in filtered_schedule.iterrows():
                    st.markdown(f"""
                    <div style="border: 2px solid #333; border-radius: 10px; padding: 15px; margin: 10px 0; background-color: #f9f9f9;">
                        <h4 style="margin: 0; color: #333;">Time: {row['Time']}</h4>
                        <p style="margin: 5px 0; color: #555;"><strong>Venue:</strong> {row['Venue']}</p>
                        <p style="margin: 5px 0; color: #555;"><strong>Subject:</strong> {row['Subject']}</p>
                        <p style="margin: 5px 0; color: #555;"><strong>Department:</strong> {row['Department']}</p>
                        <p style="margin: 5px 0; color: #555;"><strong>Class:</strong> {row['Class']}</p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.write(f"No timetable found for Department **{department}** and Class **{class_name}** on **{day}**.")
        else:
            st.write(f"Could not load schedule for **{day}**.")

    # Section for electives
    elective_input = st.text_input("Enter Electives (e.g., 'FOM 3B, FOA 3C'):").strip().upper()  # Single input field for electives

    if elective_input:  # Only proceed if input is entered
        # Process the input and split into subjects and classes
        elective_parts = [ep.strip() for ep in elective_input.split(',')]
        
        valid_electives = []
        for ep in elective_parts:
            parts = ep.split()
            if len(parts) == 2:
                subject, class_name = parts
                valid_electives.append((subject, class_name))
            else:
                st.write(f"Skipping invalid input: '{ep}'. Ensure it is in the format 'Subject Class'.")

        if valid_electives:
            schedule_df = load_sheet(day)  # Load the schedule for the specified day
            if schedule_df is not None:
                # Filter the DataFrame for the specified electives
                elective_subjects = {subject for subject, _ in valid_electives}
                elective_classes = {class_name for _, class_name in valid_electives}
                filtered_electives = schedule_df[
                    (schedule_df['Subject'].isin(elective_subjects)) & 
                    (schedule_df['Class'].isin(elective_classes))
                ]
                
                # Ensure 'Time' is sorted correctly. This assumes 'Time' is a string that represents time in a sortable format.
                filtered_electives = filtered_electives.sort_values(by='Time')
                
                if not filtered_electives.empty:
                    st.write(f"### Elective Timetable on **{day}**:")
                    
                    for index, row in filtered_electives.iterrows():
                        st.markdown(f"""
                        <div style="border: 2px solid #333; border-radius: 10px; padding: 15px; margin: 10px 0; background-color: #f9f9f9;">
                            <h4 style="margin: 0; color: #333;">Time: {row['Time']}</h4>
                            <p style="margin: 5px 0; color: #555;"><strong>Venue:</strong> {row['Venue']}</p>
                            <p style="margin: 5px 0; color: #555;"><strong>Subject:</strong> {row['Subject']}</p>
                            <p style="margin: 5px 0; color: #555;"><strong>Department:</strong> {row['Department']}</p>
                            <p style="margin: 5px 0; color: #555;"><strong>Class:</strong> {row['Class']}</p>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.write(f"No elective timetable found on **{day}**.")
            else:
                st.write(f"Could not load schedule for **{day}**.")

# Main function to run the app
def main():
    st.title("Timetable Viewer")
    show_schedule()

if __name__ == "__main__":
    main()

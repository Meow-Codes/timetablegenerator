# 🗓️ Timetable Scheduling Automation

This project automates the scheduling of courses for an academic session, targeting stakeholders such as **timetable coordinators, faculty, students, HoDs**, and **Dean Academics**.  
The software reads configuration and course data from **CSV/Excel** files and generates an **optimized timetable in Excel format**.

---

## ⚙️ Setup Instructions

### ✅ Install Dependencies

1. Ensure **Python 3.8+** is installed.
2. Install required packages:
   ```bash
   pip install -r requirements.txt

📂 Prepare Data Files
Create a data/ directory and place the following CSV files inside it:

courses.csv – Course details (e.g., course_code, department, semester, section_id, etc.)

config.csv – Scheduling parameters (e.g., slot_duration_minutes, scheduling_days, etc.)

rooms.csv – Room details (e.g., room_number, capacity, type)

sections.csv – Section details (e.g., section_id, department, semester, etc.)

faculty.csv – Faculty details (faculty_id, faculty_name, department)

assistants.csv – Assistant details (optional)

elective_enrollments.csv – Elective enrollment data (optional)

Note: Ensure the directory structure is maintained as data/ relative to the script.

▶️ Run the Script
To generate the timetable, run:

bash
Copy
Edit
python generate_timetable.py
The output will be saved in the output/ directory as:

Copy
Edit
timetable_YYYYMMDD_HHMMSS.xlsx
✅ Requirements Satisfied
The current implementation fulfills the following requirements from Timetable_Requirements-CS301 SE course.xlsx:

REQ-02-Config (Mandatory):
Reads config and course data from CSV files. Parameters such as slot_duration_minutes and scheduling_days are loaded via config_df.

REQ-03 (Mandatory):
Ensures classrooms have enough capacity. Supports splitting students into batches for labs based on sections.csv.

REQ-04-CONFLICTS (Mandatory):
Avoids scheduling multiple course components (lectures, tutorials, labs) on the same day.

REQ-05 (Mandatory):
Schedules courses with the same course code across departments separately using a unique timetable_key.

REQ-06 (Mandatory):
Adheres to the LTPSC structure. Slot durations are calculated to meet credit hour requirements.

REQ-07 (Mandatory):
Electives grouped into baskets and scheduled to avoid conflicts. Uses basket_schedules for synchronization.

REQ-08 (Mandatory):
Allocates labs based on room capacity. Batches created if enrollment exceeds lab room capacity.

REQ-09-BREAKS (Desired):
Considers morning (10:30–11:00) and staggered lunch breaks (CSE: 13:00–14:30, DSAI: 13:15–14:45, ECE: 13:30–15:00).

REQ-10-FACULTY (Mandatory):
Prevents back-to-back classes. Tries to maintain a gap (3 hrs) by controlling daily slot assignments.

REQ-14-VIEW (Mandatory):
Exports Excel timetable with sheets for faculty, students, and coordinators, including electives.

REQ-16-ANALYTICS (Desired):
Room usage statistics included in a Statistics sheet. Instructor and student effort analytics pending.

REQ-18-LUNCH (Mandatory):
Respects staggered departmental lunch breaks defined in lunch_schedule.

⚠️ Notes
❌ Unsatisfied Requirements
These are currently not implemented but can be added:

REQ-01 – Modify existing timetable

REQ-11 – Faculty preferences

REQ-12 – Reserved slots

REQ-13 – Google Calendar integration

REQ-15 – Exam timetable

REQ-17 – Teaching/Lab assistants

📸 Snapshots
Please attach screenshots of the generated Excel timetable to validate:

A section view (e.g., CSE 6A)

Elective course schedule

Statistics sheet

🚀 Future Enhancements
Faculty preference inputs

Reserved slot configurations

Exam scheduling

Google Calendar integration
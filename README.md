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

### 📂 Prepare Data Files

Create a data/ directory and place the following CSV files inside:

1. courses.csv: Course details (course_id, department, semester, course_code, course_name, lecture_hours, tutorial_hours, practical_hours, self_study_hours, credits, faculty_ids, is_elective, basket_id, combined, enrollment, section_id)

2. config.csv: Scheduling parameters (slot_duration_minutes, scheduling_days, etc.)

3. rooms.csv: Room details (room_id, room_number, capacity, type, bench_capacity)

4. sections.csv: Section details (section_id, batch_name, year, department, strength)

5. faculty.csv: Faculty details (faculty_id, faculty_name)

6. assistants.csv: Assistant details (optional) (assistant_id, assistant_name, course_id, preference_day, preference_start_time, preference_end_time, is_lab_eligible)

7. elective_enrollments.csv: Elective enrollment data (optional) (section_id, course_id, enrollment)

Note: Make sure the directory structure is maintained as data/ relative to the script.

### ▶️ Run the Script

To generate the timetable, run the following command:
```bash
python generate_timetable.py
```
The generated timetable will be saved in the output/ directory as:
```bash
timetable.html and timetable.xlsx
```

### Requirements Satisfied

The current implementation addresses the following requirements from the Timetable_Requirements-CS301 SE course.xlsx document:

1. REQ-02-Config (Mandatory): The software reads configuration parameters (e.g., slot_duration_minutes, scheduling_days) and course data (course_name, LTPSC, instructor_name, classroom details, etc.) from CSV files. Configuration is loaded via config_df from data/config.csv.

2. REQ-03 (Mandatory): Courses are scheduled in classrooms with sufficient capacity based on registered students. If a single room is insufficient, the code supports splitting students into batches (e.g., for labs), though current implementation assumes sections are predefined in sections.csv.

3. REQ-04-CONFLICTS (Mandatory): The software avoids scheduling multiple components of a course on the same day by prioritizing different days for lectures, tutorials, and practicals. Labs can follow lectures/tutorials as per the logic in scheduling loops.

4. REQ-05 (Mandatory): Courses with the same course code across different departments are scheduled separately by using unique timetable_key (dept_semester_section).

5. REQ-06 (Mandatory): The scheduling adheres to the LTPSC structure (Lecture, Tutorial, Practical, Self-study) by calculating slot durations (e.g., 3 slots for 1.5+ lecture hours, 4 slots for practicals) to meet credit requirements.

6. REQ-08 (Mandatory): Lab sessions are allocated based on lab room capacity, with batches created if the total enrollment exceeds capacity (e.g., using lab_capacity and batch calculations).

7. REQ-09-BREAKS (Desired): Morning breaks (10:30-11:00) and lunch breaks (staggered by department: CSE 13:00-14:30, DSAI 13:15-14:45, ECE 13:30-15:00) are included and respected during slot availability checks.

8. REQ-10-FACULTY (Mandatory): Consecutive classes are avoided by checking slot availability across the day. A 3-hour gap is indirectly enforced by limiting scheduling attempts per day, though explicit gap enforcement could be added.

9. REQ-18-LUNCH (Mandatory): Lunch breaks are staggered by department to avoid overcrowding, as defined in lunch_schedule.

### 🚀 Future Enhancements

Faculty preference input interface

Reserved slot configuration support

Exam timetable scheduling

Google Calendar integration
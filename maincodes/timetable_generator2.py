import pandas as pd
import random
import datetime
import copy
from collections import defaultdict
from typing import List, Dict, Tuple, Set

# Define data structures
class TimeSlot:
    def __init__(self, day: str, start_time: datetime.time, end_time: datetime.time, slot_type: str = "Regular"):
        self.day = day
        self.start_time = start_time
        self.end_time = end_time
        self.slot_type = slot_type

    def duration_hours(self) -> float:
        start_dt = datetime.datetime.combine(datetime.date.today(), self.start_time)
        end_dt = datetime.datetime.combine(datetime.date.today(), self.end_time)
        duration = end_dt - start_dt
        return duration.total_seconds() / 3600

    def __str__(self) -> str:
        return f"{self.day} {self.start_time.strftime('%H:%M')}-{self.end_time.strftime('%H:%M')}"

    def overlaps(self, other) -> bool:
        if self.day != other.day:
            return False
        return self.start_time < other.end_time and self.end_time > other.start_time

class Room:
    def __init__(self, room_id: str, room_type: str, capacity: int):
        self.room_id = room_id
        self.room_type = room_type
        self.capacity = capacity

    def __str__(self) -> str:
        return f"{self.room_id} ({self.room_type}, capacity: {self.capacity})"

class Course:
    def __init__(self, course_id: str, course_name: str, professor_id: str, total_students: int,
                 num_lectures: int, num_labs: int, num_tutorials: int, fixed_classroom: str = None):
        self.course_id = course_id
        self.course_name = course_name
        self.professor_id = professor_id
        self.total_students = total_students
        self.num_lectures = num_lectures
        self.num_labs = num_labs
        self.num_tutorials = num_tutorials
        self.fixed_classroom = fixed_classroom

    def total_sessions(self) -> int:
        return self.num_lectures + self.num_labs + self.num_tutorials

    def __str__(self) -> str:
        return f"{self.course_id}: {self.course_name}, Students: {self.total_students}"

class Session:
    def __init__(self, course: Course, session_type: str, room: Room, time_slot: TimeSlot, session_id: int):
        self.course = course
        self.session_type = session_type
        self.room = room
        self.time_slot = time_slot
        self.session_id = session_id

    def __str__(self) -> str:
        return f"{self.course.course_id} ({self.session_type}, {self.course.professor_id}, {self.room.room_id}) [ID:{self.session_id}]"

class Timetable:
    def __init__(self):
        self.sessions = []
        self.course_day_sessions = defaultdict(lambda: defaultdict(list))
        self.room_timeslot_map = {}
        self.professor_timeslot_map = {}
        self.lab_days = defaultdict(set)

    def add_session(self, session: Session) -> None:
        self.sessions.append(session)
        self.course_day_sessions[session.course.course_id][session.time_slot.day].append(session)
        key = (session.room.room_id, session.time_slot)
        self.room_timeslot_map[key] = session
        prof_key = (session.course.professor_id, session.time_slot)
        self.professor_timeslot_map[prof_key] = session
        if session.session_type == "Lab":
            self.lab_days[session.time_slot.day].add(session.course.course_id)

    def remove_last_session(self) -> Session:
        if not self.sessions:
            return None
        session = self.sessions.pop()
        self.course_day_sessions[session.course.course_id][session.time_slot.day].remove(session)
        if not self.course_day_sessions[session.course.course_id][session.time_slot.day]:
            del self.course_day_sessions[session.course.course_id][session.time_slot.day]
        key = (session.room.room_id, session.time_slot)
        if key in self.room_timeslot_map:
            del self.room_timeslot_map[key]
        prof_key = (session.course.professor_id, session.time_slot)
        if prof_key in self.professor_timeslot_map:
            del self.professor_timeslot_map[prof_key]
        if session.session_type == "Lab":
            self.lab_days[session.time_slot.day].discard(session.course.course_id)
        return session

    def is_room_available(self, room: Room, time_slot: TimeSlot) -> bool:
        if (room.room_id, time_slot) in self.room_timeslot_map:
            return False
        for slot in [ts for r, ts in self.room_timeslot_map.keys() if r == room.room_id]:
            if time_slot.overlaps(slot):
                return False
        return True

    def is_professor_available(self, professor_id: str, time_slot: TimeSlot) -> bool:
        if (professor_id, time_slot) in self.professor_timeslot_map:
            return False
        for slot in [ts for p, ts in self.professor_timeslot_map.keys() if p == professor_id]:
            if time_slot.overlaps(slot):
                return False
        return True

    def count_session_type_on_day(self, course_id: str, session_type: str, day: str) -> int:
        count = 0
        for session in self.course_day_sessions[course_id].get(day, []):
            if session.session_type == session_type:
                count += 1
        return count

    def count_sessions_on_day(self, course_id: str, day: str) -> int:
        return len(self.course_day_sessions[course_id].get(day, []))

class TimetableGenerator:
    def __init__(self):
        self.working_days = ["MON", "TUE", "WED", "THU", "FRI"]
        self.fixed_break_slots = [
            ("Morning Break", datetime.time(10, 30), datetime.time(11, 0)),
            ("Lunch", datetime.time(13, 30), datetime.time(14, 30))
        ]
        self.optional_snack_slot = ("Snacks", datetime.time(16, 30), datetime.time(17, 0))
        self.working_hours = {"start": datetime.time(9, 0), "end": datetime.time(17, 0)}
        self.durations = {"Lecture": 1.5, "Lab": 2.0, "Tutorial": 1.0}
        self.max_backtrack_attempts = 2000  # Increased for better scheduling
        self.backtrack_count = 0
        self.session_ids = defaultdict(lambda: {"Lecture": 0, "Tutorial": 0, "Lab": 0})

    def generate_time_slots(self) -> List[TimeSlot]:
        time_slots = []
        for day in self.working_days:
            for break_name, start, end in self.fixed_break_slots:
                time_slots.append(TimeSlot(day, start, end, break_name))
            current_time = self.working_hours["start"]
            while current_time < self.working_hours["end"]:
                for session_type, duration in self.durations.items():
                    duration_delta = datetime.timedelta(hours=duration)
                    end_time_dt = datetime.datetime.combine(datetime.date.today(), current_time) + duration_delta
                    end_time = end_time_dt.time()
                    if end_time <= self.working_hours["end"]:
                        slot = TimeSlot(day, current_time, end_time)
                        valid = True
                        for _, break_start, break_end in self.fixed_break_slots:
                            break_slot = TimeSlot(day, break_start, break_end)
                            if slot.overlaps(break_slot):
                                valid = False
                                break
                        if valid:
                            time_slots.append(slot)
                current_dt = datetime.datetime.combine(datetime.date.today(), current_time)
                current_dt += datetime.timedelta(minutes=30)
                current_time = current_dt.time()
        return time_slots

    def filter_time_slots(self, time_slots: List[TimeSlot], session_type: str) -> List[TimeSlot]:
        duration = self.durations[session_type]
        return [slot for slot in time_slots if slot.slot_type == "Regular" and abs(slot.duration_hours() - duration) < 0.01]

    def select_best_room(self, course: Course, session_type: str, available_rooms: List[Room]) -> Room:
        if not available_rooms:
            return None
        if course.fixed_classroom and session_type != "Lab":
            for room in available_rooms:
                if room.room_id == course.fixed_classroom:
                    return room
        if session_type == "Lab":
            lab_rooms = [room for room in available_rooms if room.room_type == "LabRoom"]
            if lab_rooms:
                suitable_labs = [r for r in lab_rooms if r.capacity >= course.total_students]
                return min(suitable_labs, key=lambda r: r.capacity) if suitable_labs else max(lab_rooms, key=lambda r: r.capacity)
        classrooms = [room for room in available_rooms if room.room_type == "Classroom"]
        if classrooms:
            suitable_classrooms = [r for r in classrooms if r.capacity >= course.total_students]
            return min(suitable_classrooms, key=lambda r: r.capacity) if suitable_classrooms else max(classrooms, key=lambda r: r.capacity)
        return min(available_rooms, key=lambda r: abs(r.capacity - course.total_students))

    def get_available_rooms(self, rooms: List[Room], course: Course, session_type: str, time_slot: TimeSlot, timetable: Timetable) -> List[Room]:
        available_rooms = []
        for room in rooms:
            if not timetable.is_room_available(room, time_slot):
                continue
            if session_type == "Lab" and room.room_type != "LabRoom":
                continue
            if course.fixed_classroom and session_type != "Lab" and room.room_id != course.fixed_classroom:
                continue
            available_rooms.append(room)
        return available_rooms

    def assign_session(self, course: Course, session_type: str, available_time_slots: List[TimeSlot], rooms: List[Room], timetable: Timetable) -> bool:
        valid_time_slots = self.filter_time_slots(available_time_slots, session_type)
        random.shuffle(valid_time_slots)
        for time_slot in valid_time_slots:
            if not timetable.is_professor_available(course.professor_id, time_slot):
                continue
            if session_type == "Lab":
                day_idx = self.working_days.index(time_slot.day)
                prev_day = self.working_days[day_idx - 1] if day_idx > 0 else None
                next_day = self.working_days[day_idx + 1] if day_idx < len(self.working_days) - 1 else None
                if (prev_day and course.course_id in timetable.lab_days[prev_day]) or (next_day and course.course_id in timetable.lab_days[next_day]):
                    continue
            available_rooms = self.get_available_rooms(rooms, course, session_type, time_slot, timetable)
            if not available_rooms:
                continue
            best_room = self.select_best_room(course, session_type, available_rooms)
            if best_room:
                session_id = self.session_ids[course.course_id][session_type]
                session = Session(course, session_type, best_room, time_slot, session_id)
                timetable.add_session(session)
                self.session_ids[course.course_id][session_type] += 1
                available_time_slots.remove(time_slot)
                print(f"Scheduled {session_type} [ID:{session_id}] for {course.course_id} on {time_slot}")
                return True
        print(f"Failed to schedule {session_type} for {course.course_id}")
        return False

    def generate_timetable(self, courses: List[Course], rooms: List[Room]) -> Timetable:
        timetable = Timetable()
        all_time_slots = self.generate_time_slots()
        # Prioritize courses with labs to ensure lab days are available
        sorted_courses = sorted(courses, key=lambda c: (-c.num_labs, -c.total_sessions(), -c.total_students))
        assignment_stack = []
        for course in sorted_courses:
            available_time_slots = copy.deepcopy(all_time_slots)
            # Schedule labs first to ensure lab day constraints are met
            for session_type, count in [("Lab", course.num_labs), ("Lecture", course.num_lectures), ("Tutorial", course.num_tutorials)]:
                for _ in range(count):
                    successful = self.assign_session(course, session_type, available_time_slots, rooms, timetable)
                    if successful:
                        assignment_stack.append((course, session_type))
                    else:
                        return self.backtrack(timetable, courses, rooms, assignment_stack)
        return timetable

    def backtrack(self, timetable: Timetable, courses: List[Course], rooms: List[Room], assignment_stack: List[Tuple]) -> Timetable:
        self.backtrack_count += 1
        if self.backtrack_count > self.max_backtrack_attempts:
            print(f"Warning: Maximum backtracking attempts ({self.max_backtrack_attempts}) reached.")
            return timetable
        session = timetable.remove_last_session()
        if not assignment_stack:
            print("Error: No more assignments to backtrack.")
            return timetable
        last_course, last_session_type = assignment_stack.pop()
        all_time_slots = self.generate_time_slots()
        successful = self.assign_session(last_course, last_session_type, all_time_slots, rooms, timetable)
        if successful:
            assignment_stack.append((last_course, last_session_type))
            return self.generate_timetable(courses, rooms)
        return self.backtrack(timetable, courses, rooms, assignment_stack)

    def validate_timetable(self, timetable: Timetable, courses: List[Course], rooms: List[Room]):
        all_time_slots = self.generate_time_slots()
        for course in courses:
            expected_lectures = course.num_lectures
            expected_tutorials = course.num_tutorials
            expected_labs = course.num_labs
            assigned_lectures = sum(1 for s in timetable.sessions if s.course.course_id == course.course_id and s.session_type == "Lecture")
            assigned_tutorials = sum(1 for s in timetable.sessions if s.course.course_id == course.course_id and s.session_type == "Tutorial")
            assigned_labs = sum(1 for s in timetable.sessions if s.course.course_id == course.course_id and s.session_type == "Lab")
            if assigned_lectures != expected_lectures or assigned_tutorials != expected_tutorials or assigned_labs != expected_labs:
                print(f"Validation Failed for {course.course_id}: Expected L:{expected_lectures}, T:{expected_tutorials}, P:{expected_labs}, "
                      f"Got L:{assigned_lectures}, T:{assigned_tutorials}, P:{assigned_labs}")
                # Reschedule missing sessions
                available_time_slots = copy.deepcopy(all_time_slots)
                for session_type, expected, assigned in [("Lecture", expected_lectures, assigned_lectures),
                                                        ("Tutorial", expected_tutorials, assigned_tutorials),
                                                        ("Lab", expected_labs, assigned_labs)]:
                    for _ in range(expected - assigned):
                        self.assign_session(course, session_type, available_time_slots, rooms, timetable)

def main():
    # Read the CSV file
    csv_file = input("Enter the CSV file path (e.g., Data/CSE/4th_sem_A.csv): ")
    df = pd.read_csv(csv_file)
    df.columns = ["Course Code", "Course Name", "L", "T", "P", "S", "C", "Faculty", "Classroom"]

    # Verify data and parse correctly
    print("Loaded CSV Data:")
    print(df)
    response = input("Is the data above correct? (yes/no): ")
    if response.lower() != "yes":
        print("Please correct the CSV file and try again.")
        exit()

    # Create rooms
    rooms = []
    for classroom in df["Classroom"].unique():
        if pd.notna(classroom):
            rooms.append(Room(classroom, "Classroom", 60))
    rooms.extend([
        Room("L107", "LabRoom", 40),
        Room("L106", "LabRoom", 40)
    ])

    # Create courses with corrected parsing
    courses = []
    for _, row in df.iterrows():
        try:
            L = int(float(row["L"]))  # Ensure L is an integer
            T = int(float(row["T"]))  # Ensure T is an integer
            P = int(float(row["P"]))  # Ensure P is an integer
        except (ValueError, TypeError) as e:
            print(f"Error parsing L, T, P for {row['Course Code']}: L={row['L']}, T={row['T']}, P={row['P']}. Error: {e}")
            continue
        num_lectures = int(L / 1.5)
        num_tutorials = T
        num_labs = P // 2
        print(f"Parsed {row['Course Code']}: L={L}, T={T}, P={P} -> Lectures={num_lectures}, Tutorials={num_tutorials}, Labs={num_labs}")
        courses.append(Course(
            course_id=row["Course Code"],
            course_name=row["Course Name"],
            professor_id=row["Faculty"],
            total_students=60,
            num_lectures=num_lectures,
            num_labs=num_labs,
            num_tutorials=num_tutorials,
            fixed_classroom=row["Classroom"] if pd.notna(row["Classroom"]) else None
        ))

    # Initialize timetable generator
    generator = TimetableGenerator()
    print("Generating timetable...")
    timetable = generator.generate_timetable(courses, rooms)
    generator.validate_timetable(timetable, courses, rooms)

    # Add optional snacks
    snack_slot = TimeSlot("", generator.optional_snack_slot[1], generator.optional_snack_slot[2], "Snacks")
    for day in generator.working_days:
        slot = TimeSlot(day, generator.optional_snack_slot[1], generator.optional_snack_slot[2], "Snacks")
        if all(not timetable.is_room_available(Room("Dummy", "Classroom", 60), TimeSlot(day, slot.start_time, slot.end_time)) for s in timetable.sessions if s.time_slot.day == day and s.time_slot.overlaps(slot)):
            dummy_room = Room("Cafeteria", "Classroom", 100)
            dummy_course = Course("Snacks", "Snacks", "N/A", 0, 0, 0, 0)
            timetable.add_session(Session(dummy_course, "Snacks", dummy_room, slot, 0))

    # Prepare timetable for HTML export
    base_slots = [f"{h:02d}:{m:02d}-{h + (m + 30) // 60:02d}:{(m + 30) % 60:02d}" for h in range(9, 17) for m in [0, 30]]
    slot_count = len(base_slots)
    timetable_grid = {day: ["" for _ in range(slot_count)] for day in generator.working_days}
    for day in generator.working_days:
        for i in range(slot_count):
            start_str, end_str = base_slots[i].split("-")
            start_time = datetime.datetime.strptime(start_str, "%H:%M").time()
            end_time = datetime.datetime.strptime(end_str, "%H:%M").time()
            slot = TimeSlot(day, start_time, end_time)
            for session in timetable.sessions:
                if session.time_slot.day == day and session.time_slot.overlaps(slot):
                    timetable_grid[day][i] = str(session)
            for break_name, break_start, break_end in generator.fixed_break_slots:
                break_slot = TimeSlot(day, break_start, break_end)
                if slot.overlaps(break_slot):
                    timetable_grid[day][i] = break_name
            snack_slot = TimeSlot(day, generator.optional_snack_slot[1], generator.optional_snack_slot[2])
            if slot.overlaps(snack_slot) and any(s.session_type == "Snacks" and s.time_slot.day == day for s in timetable.sessions):
                timetable_grid[day][i] = "Snacks"

    # Define a list of distinct colors
    distinct_colors = ["#FF6347", "#4682B4", "#32CD32", "#FFD700", "#6A5ACD", "#FF69B4", "#00CED1", "#FFA500", "#20B2AA", "#DAA520"]
    course_codes = df["Course Code"].unique()
    course_colors = {course: distinct_colors[i % len(distinct_colors)] for i, course in enumerate(course_codes)}

    # Export to HTML
    html_content = r"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>IIIT Dharwad Timetable - Semester IV (Dec 2024 - Apr 2025)</title>
    <style>
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid black; padding: 8px; text-align: center; }
        th { background-color: #f2f2f2; }
        .break { background-color: #ffa07a; }
        .lunch { background-color: #d3d3d3; }
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>
<body>
    <h2>IIIT Dharwad Timetable - Semester IV (Dec 2024 - Apr 2025)</h2>
    <h3>Section A - Roll No 23BCS001 to 23BCS070</h3>
    <button onclick="exportToExcel()">Export to Excel</button>
    <table id="timetable">
        <tr>
            <th>Time/Day</th>
"""
    for slot in base_slots:
        html_content += f"<th>{slot}</th>"
    html_content += "</tr>"
    for day in generator.working_days:
        html_content += f"<tr><td>{day}</td>"
        for i, slot_type in enumerate(base_slots):
            cell_content = timetable_grid[day][i]
            if cell_content in ["Morning Break", "Lunch", "Snacks"]:
                class_name = {"Morning Break": "break", "Lunch": "lunch", "Snacks": "break"}.get(cell_content, "")
                html_content += f"<td class='{class_name}'>{cell_content}</td>"
            else:
                course_match = next((c for c in course_codes if c in cell_content), None)
                color = course_colors.get(course_match, "#FFFFFF")
                html_content += f"<td style='background-color: {color};'>{cell_content}</td>"
        html_content += "</tr>"
    html_content += r"""
    </table>
    <script>
        function exportToExcel() {
            var table = document.getElementById("timetable");
            var wb = XLSX.utils.book_new();
            var ws_data = [];
            var colors = [];
            for (var i = 0; i < table.rows.length; i++) {
                var row = [], row_colors = [];
                for (var j = 0; j < table.rows[i].cells.length; j++) {
                    row.push(table.rows[i].cells[j].innerText);
                    var bgColor = table.rows[i].cells[j].style.backgroundColor;
                    if (bgColor && bgColor.startsWith('rgb')) {
                        var rgb = bgColor.match(/\d+/g);
                        bgColor = "#" + ((1 << 24) + (parseInt(rgb[0]) << 16) + (parseInt(rgb[1]) << 8) + parseInt(rgb[2])).toString(16).slice(1).toUpperCase();
                    }
                    row_colors.push(bgColor || "");
                }
                ws_data.push(row);
                colors.push(row_colors);
            }
            var ws = XLSX.utils.aoa_to_sheet(ws_data);
            for (var r = 0; r < ws_data.length; r++) {
                for (var c = 0; c < ws_data[r].length; c++) {
                    var cell_ref = XLSX.utils.encode_cell({r: r, c: c});
                    if (!ws[cell_ref]) ws[cell_ref] = {};
                    if (colors[r][c]) {
                        ws[cell_ref].s = {fill: {fgColor: {rgb: colors[r][c].replace("#", "")}}};
                    }
                }
            }
            XLSX.utils.book_append_sheet(wb, ws, "Timetable");
            XLSX.writeFile(wb, "timetable.xlsx");
        }
    </script>
</body>
</html>
"""

    with open("timetable.html", "w") as f:
        f.write(html_content)

    print(f"Timetable generated successfully! Check timetable.html and use the 'Export to Excel' button to download the timetable as an Excel file with colors.")

if __name__ == "__main__":
    random.seed(42)
    main()
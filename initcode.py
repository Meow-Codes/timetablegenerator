import csv
from collections import defaultdict
import random

class TimetableGenerator:
    def __init__(self, teachers, subjects, rooms, time_slots, constraints=None):
        self.teachers = teachers
        self.subjects = subjects
        self.rooms = rooms
        self.time_slots = time_slots  # Dict with course types and durations
        self.constraints = constraints if constraints else {}
        self.timetable = defaultdict(lambda: defaultdict(dict))

    def is_valid_assignment(self, teacher, subject, room, time_slot):
        """Check if the assignment satisfies all constraints."""
        if teacher not in self.teachers or subject not in self.subjects or room not in self.rooms:
            return False
        
        # Check time-slot constraints
        if time_slot in self.timetable and room in self.timetable[time_slot]:
            return False  # Room is already occupied

        # Check teacher's availability
        for ts, assignments in self.timetable.items():
            if teacher in [details['teacher'] for details in assignments.values()]:
                if ts == time_slot:
                    return False  # Teacher is occupied

        return True

    def generate(self):
        """Generate the timetable based on the provided constraints."""
        for subject, details in self.subjects.items():
            course_type = details.get('type', 'lecture')  # Default to lecture
            duration = self.time_slots[course_type]

            teacher = random.choice(self.teachers)
            room = random.choice(self.rooms)
            time_slot = random.choice(self.time_slots[course_type])

            while not self.is_valid_assignment(teacher, subject, room, time_slot):
                teacher = random.choice(self.teachers)
                room = random.choice(self.rooms)
                time_slot = random.choice(self.time_slots[course_type])

            self.timetable[time_slot][room] = {
                "teacher": teacher,
                "subject": subject,
                "type": course_type,
                "duration": duration
            }

        return self.timetable

    def display(self):
        """Display the generated timetable."""
        print("Timetable:")
        for time_slot, rooms in self.timetable.items():
            print(f"Time Slot: {time_slot}")
            for room, details in rooms.items():
                print(f"  Room: {room} | Teacher: {details['teacher']} | Subject: {details['subject']} | Type: {details['type']} | Duration: {details['duration']} hrs")

# Sample Data
teachers = ["Mr. Smith", "Ms. Johnson", "Dr. Brown"]
subjects = {
    "Math": {"type": "lecture"},
    "Science": {"type": "lab"},
    "History": {"type": "tutorial"},
    "English": {"type": "lecture"}
}
rooms = ["Room A", "Room B", "Room C"]
time_slots = {
    "lecture": ["9:00-10:30", "11:00-12:30", "2:00-3:30"],
    "tutorial": ["10:30-11:30", "1:00-2:00", "3:30-4:30"],
    "lab": ["9:00-11:00", "1:00-3:00"]
}

# Create and Generate Timetable
generator = TimetableGenerator(teachers, subjects, rooms, time_slots)
timetable = generator.generate()

generator.display()

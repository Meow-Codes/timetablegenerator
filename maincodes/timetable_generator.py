import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import random

# Color mapping for courses (matching the image)
COLOR_MAPPING = {
    'CS206': '#ADD8E6',  # Light Blue
    'MA202': '#FFD580',  # Light Orange
    'CS301': '#FFB6C1',  # Pink
    'CS204': '#90EE90',  # Light Green
    'CS310': '#E6E6FA',  # Light Purple
    'HS205': '#DEB887'   # Light Brown
}

class TimetableGenerator:
    def __init__(self, csv_path):
        self.csv_path = csv_path
        self.courses_df = None
        self.timetable = None
        
        # Define fixed time slots as per requirements
        self.time_slots = [
            '09:00-10:30',  # Morning lecture slot
            '11:00-12:30',  # Post-break lecture slot
            '12:30-13:30',  # Tutorial slot
            '14:30-16:30',  # Lab slot
            '16:30-17:00',  # Tutorial slot
            '17:00-18:30'   # Evening lecture slot
        ]
        
        self.breaks = {
            'Morning Break': '10:30-11:00',
            'Lunch Break': '13:30-14:30',
            'Evening Break': '16:15-17:15'
        }
        
        self.days = ['MON', 'TUE', 'WED', 'THU', 'FRI']
        
        # Color scheme for courses
        self.color_scheme = {
            'CS206': '#ADD8E6',  # Light Blue
            'MA202': '#FFD580',  # Light Orange
            'CS301': '#FFB6C1',  # Pink
            'CS204': '#90EE90',  # Light Green
            'CS310': '#E6E6FA',  # Light Purple
            'HS205': '#DEB887'   # Light Brown
        }
        
        # Available rooms
        self.classrooms = ['C104', 'C105', 'C106']
        self.labs = ['L106', 'L107']
        
    def read_course_data(self):
        """Read and validate CSV file"""
        try:
            self.courses_df = pd.read_csv(self.csv_path)
            self.validate_csv()
            print("Courses loaded successfully:")
            print(self.courses_df[['Course Code', 'Course Name', 'L', 'T', 'P']])
            
            # Calculate actual sessions needed
            self.courses_df['Lecture_Sessions'] = (self.courses_df['L'] / 1.5).apply(np.ceil)
            self.courses_df['Tutorial_Sessions'] = self.courses_df['T']
            self.courses_df['Lab_Sessions'] = (self.courses_df['P'] / 2).apply(np.ceil)
            
            return self.courses_df
        except Exception as e:
            print(f"Error reading CSV: {str(e)}")
            return None
    
    def validate_csv(self):
        """Validate CSV file structure"""
        required_columns = ['Course Code', 'Course Name', 'L', 'T', 'P', 'S', 'C', 'Faculty']
        df_columns = self.courses_df.columns.tolist()
        
        missing_columns = [col for col in required_columns if col not in df_columns]
        if missing_columns:
            raise ValueError(f"Missing required columns: {missing_columns}")
            
        # Validate numeric columns
        numeric_columns = ['L', 'T', 'P', 'S', 'C']
        for col in numeric_columns:
            if not pd.to_numeric(self.courses_df[col], errors='coerce').notnull().all():
                raise ValueError(f"Column {col} contains non-numeric values")
    
    def is_slot_available(self, day, time_slot):
        """Check if a time slot is available"""
        return pd.isna(self.timetable.at[time_slot, day])
    
    def is_valid_lecture_slot(self, time_slot):
        """Check if slot is valid for lectures (1.5 hours)"""
        return time_slot in ['09:00-10:30', '11:00-12:30', '17:00-18:30']
    
    def is_valid_tutorial_slot(self, time_slot):
        """Check if slot is valid for tutorials (1 hour)"""
        return time_slot in ['12:30-13:30', '16:30-17:00']
    
    def is_valid_lab_slot(self, time_slot):
        """Check if slot is valid for labs (2 hours)"""
        return time_slot == '14:30-16:30'
    
    def generate_timetable(self):
        """Generate weekly timetable based on course data"""
        if self.courses_df is None:
            print("No course data available")
            return None
            
        print("Initializing timetable...")
        self.timetable = pd.DataFrame(index=self.time_slots, columns=self.days)
        
        # First schedule labs as they have specific requirements
        print("Scheduling labs...")
        lab_courses = self.courses_df[self.courses_df['Lab_Sessions'] > 0]
        for _, course in lab_courses.iterrows():
            self._schedule_labs(course)
        
        # Then schedule lectures
        print("Scheduling lectures...")
        for _, course in self.courses_df.iterrows():
            self._schedule_lectures(course)
        
        # Finally schedule tutorials
        print("Scheduling tutorials...")
        for _, course in self.courses_df.iterrows():
            self._schedule_tutorials(course)
        
        print("Timetable generation completed!")
        return self.timetable
    
    def _schedule_lectures(self, course):
        """Schedule lectures (1.5 hour sessions)"""
        sessions_needed = int(course['Lecture_Sessions'])
        course_code = course['Course Code']
        max_attempts = 50
        
        while sessions_needed > 0 and max_attempts > 0:
            day = random.choice(self.days)
            slot = random.choice([s for s in self.time_slots if self.is_valid_lecture_slot(s)])
            
            if self.is_slot_available(day, slot):
                classroom = random.choice(self.classrooms)
                self.timetable.at[slot, day] = f"{course_code} (L) - {classroom}"
                sessions_needed -= 1
            max_attempts -= 1
        
        if sessions_needed > 0:
            print(f"Warning: Could not schedule all lectures for {course_code}")
    
    def _schedule_tutorials(self, course):
        """Schedule tutorials (1 hour sessions)"""
        sessions_needed = int(course['Tutorial_Sessions'])
        course_code = course['Course Code']
        max_attempts = 50
        
        while sessions_needed > 0 and max_attempts > 0:
            day = random.choice(self.days)
            slot = random.choice([s for s in self.time_slots if self.is_valid_tutorial_slot(s)])
            
            if self.is_slot_available(day, slot):
                self.timetable.at[slot, day] = f"{course_code} (T)"
                sessions_needed -= 1
            max_attempts -= 1
        
        if sessions_needed > 0:
            print(f"Warning: Could not schedule all tutorials for {course_code}")
    
    def _schedule_labs(self, course):
        """Schedule lab sessions (2 hour sessions)"""
        sessions_needed = int(course['Lab_Sessions'])
        course_code = course['Course Code']
        
        for _ in range(sessions_needed):
            scheduled = False
            for day in random.sample(self.days, len(self.days)):
                slot = '14:30-16:30'  # Labs are always 2 hours
                
                if self.is_slot_available(day, slot):
                    lab1, lab2 = random.sample(self.labs, 2)
                    self.timetable.at[slot, day] = (
                        f"{course_code} Lab - Batch A1 ({lab1}) / "
                        f"Batch A2 ({lab2})"
                    )
                    scheduled = True
                    break
            
            if not scheduled:
                print(f"Warning: Could not schedule lab for {course_code}")
    
    def format_timetable(self):
        """Format timetable with colors and styling"""
        def apply_style(val):
            if pd.isna(val):
                return 'background-color: white'
            
            for course_code, color in self.color_scheme.items():
                if course_code in str(val):
                    return f'background-color: {color}'
            
            if '(T)' in str(val):
                return 'background-color: #D3D3D3'  # Light gray for tutorials
            if 'Lab' in str(val):
                return 'background-color: #F0E68C'  # Light yellow for labs
            
            return 'background-color: white'
        
        styled = self.timetable.style.map(apply_style)
        styled.set_properties(**{
            'border': '1px solid black',
            'text-align': 'center',
            'padding': '8px',
            'font-size': '12px'
        })
        
        # Add custom CSS for better HTML output
        styled.set_table_styles([
            {'selector': 'th', 'props': [
                ('background-color', '#f2f2f2'),
                ('font-weight', 'bold'),
                ('border', '1px solid black')
            ]},
            {'selector': 'td', 'props': [
                ('border', '1px solid black')
            ]}
        ])
        
        return styled

def main():
    try:
        csv_path = r'C:\Users\bargh\OneDrive\Desktop\TTT\timetablegenerator\Data\CSE\4th_sem_A.csv'
        
        if not os.path.exists(csv_path):
            print(f"Error: File not found at {csv_path}")
            return
            
        print("Starting timetable generation...")
        generator = TimetableGenerator(csv_path)
        courses = generator.read_course_data()
        
        if courses is not None:
            print("\nGenerating timetable...")
            timetable = generator.generate_timetable()
            
            if timetable is not None:
                print("\nFormatting timetable...")
                styled_timetable = generator.format_timetable()
                
                print("\nSaving to HTML...")
                # Save to HTML with improved styling
                html_content = styled_timetable.to_html()
                with open('timetable.html', 'w') as f:
                    f.write('''
                        <html>
                        <head>
                            <style>
                                table { border-collapse: collapse; width: 100%; }
                                th, td { padding: 8px; border: 1px solid black; }
                                th { background-color: #f2f2f2; }
                            </style>
                        </head>
                        <body>
                    ''')
                    f.write(html_content)
                    f.write('</body></html>')
                
                print("\nTimetable has been saved to 'timetable.html'")
                print("\nWeekly Timetable Preview:")
                print(timetable)
                
                # Open the HTML file in the default browser
                import webbrowser
                webbrowser.open('timetable.html')
            else:
                print("Failed to generate timetable")
            
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main() 
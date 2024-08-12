from kivymd.app import MDApp
from kivy.uix.screenmanager import Screen
from kivy.lang import Builder
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.boxlayout import BoxLayout
from kivymd.uix.button import MDButton
import pandas as pd
import os
from datetime import datetime

KV = '''
MDScreen:
    BoxLayout:
        orientation: 'vertical'
        MDTopAppBar:
            title: "Student Attendance"
            elevation: 10
        Image:
            id: college_logo
            source: "assets/logo.png"
            size_hint: (1, 0.2)
            pos_hint: {'center_x': 0.5}
        BoxLayout:
            orientation: 'horizontal'
            size_hint: (1, 0.1)
            padding: [10, 10]
            MDLabel:
                text: "Total Students:"
                size_hint_x: 0.3
                theme_text_color: "Primary"
            TextInput:
                id: total_students
                size_hint_x: 0.7
                multiline: False
                input_filter: 'int'
                foreground_color: [0, 0, 0, 1]
        BoxLayout:
            orientation: 'horizontal'
            size_hint: (1, 0.1)
            padding: [10, 10]
            MDLabel:
                text: "Class Division:"
                size_hint_x: 0.3
                theme_text_color: "Primary"
            TextInput:
                id: class_division
                size_hint_x: 0.7
                multiline: False
                foreground_color: [0, 0, 0, 1]
        BoxLayout:
            orientation: 'horizontal'
            size_hint: (1, 0.1)
            padding: [10, 10]
            MDButton:
                text: "Start Attendance"
                on_release: app.start_attendance()
                text_color: [1, 1, 1, 1]
                md_bg_color: [0.2, 0.6, 1, 1]
        ScrollView:
            id: scroll_view
            BoxLayout:
                id: student_list
                orientation: 'vertical'
                size_hint_y: None
                height: self.minimum_height
        MDButton:
            text: "Export to Excel"
            pos_hint: {'center_x': 0.5}
            size_hint_y: 0.1
            on_release: app.export_to_excel()
            text_color: [1, 1, 1, 1]
            md_bg_color: [0.2, 0.6, 1, 1]
'''

class MainApp(MDApp):
    def __init__(self):
        super().__init__()
        self.kvs = Builder.load_string(KV)
        self.attendance_data = {}

    def build(self):
        self.theme_cls.theme_style = "Light"  # Set the theme to Light
        self.theme_cls.primary_palette = "Blue"  # Set the primary color to Blue
        return self.kvs

    def start_attendance(self):
        total_students = self.kvs.ids.total_students.text
        class_division = self.kvs.ids.class_division.text
        if not total_students.isdigit():
            print("Please enter a valid number of students.")
            return
        if not class_division:
            print("Please enter a valid class division.")
            return

        self.attendance_data = {}
        self.kvs.ids.student_list.clear_widgets()
        for i in range(1, int(total_students) + 1):
            student_layout = BoxLayout(size_hint_y=None, height=40, padding=[10, 0])
            student_label = Label(text=f"Student {i}", size_hint_x=0.6, color=[0, 0, 0, 1])
            present_button = MDButton(text="Present", on_release=lambda x, i=i: self.mark_attendance(i, "Present"), text_color=[1, 1, 1, 1], md_bg_color=[0.2, 0.6, 1, 1])
            absent_button = MDButton(text="Absent", on_release=lambda x, i=i: self.mark_attendance(i, "Absent"), text_color=[1, 1, 1, 1], md_bg_color=[1, 0.2, 0.2, 1])
            student_layout.add_widget(student_label)
            student_layout.add_widget(present_button)
            student_layout.add_widget(absent_button)
            self.kvs.ids.student_list.add_widget(student_layout)
        self.class_division = class_division

    def mark_attendance(self, student_number, status):
        self.attendance_data[student_number] = status
        print(f"Student {student_number} marked as {status}")

    def export_to_excel(self):
        # Prepare the DataFrame
        df = pd.DataFrame(list(self.attendance_data.items()), columns=["Student Number", datetime.now().strftime("%Y-%m-%d")])
        df.set_index("Student Number", inplace=True)

        # Directory for exporting files
        if not os.path.exists("exports"):
            os.makedirs("exports")

        file_path = f"exports/attendance_{self.class_division}.xlsx"

        # Check if the file exists
        if os.path.exists(file_path):
            # Load the existing file and merge with the new data
            existing_df = pd.read_excel(file_path, index_col="Student Number")
            df = pd.concat([existing_df, df], axis=1, sort=False)
        
        # Save the updated dataframe back to the Excel file
        df.to_excel(file_path)
        print(f"Attendance data has been updated in {file_path}")

if __name__ == "__main__":
    MainApp().run()

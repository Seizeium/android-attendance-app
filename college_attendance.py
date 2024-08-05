from kivymd.app import MDApp
from kivy.uix.screenmanager import Screen
from kivy.lang import Builder
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.boxlayout import BoxLayout
from kivymd.uix.button import MDButton
from kivymd.uix.snackbar import snackbar
import pandas as pd
import os

KV = '''
MDScreen:
    BoxLayout:
        orientation: 'vertical'
        MDTopAppBar:
            title: "Student Attendance"
            elevation: 10
        Image:
            id: college_logo
            source: "assets/logo.jpeg"
            size_hint: (1, 0.2)
            pos_hint: {'center_x': 0.5}
        BoxLayout:
            orientation: 'horizontal'
            size_hint: (1, 0.1)
            padding: [10, 10]
            MDLabel:
                text: "Total Students:"
                size_hint_x: 0.3
            TextInput:
                id: total_students
                size_hint_x: 0.7
                multiline: False
                input_filter: 'int'
        BoxLayout:
            orientation: 'horizontal'
            size_hint: (1, 0.1)
            padding: [10, 10]
            MDButton:
                text: "Start Attendance"
                on_release: app.start_attendance()
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
'''

class MainApp(MDApp):
    def __init__(self):
        super().__init__()
        self.kvs = Builder.load_string(KV)
        self.attendance_data = []

    def build(self):
        return self.kvs

    def start_attendance(self):
        total_students = self.kvs.ids.total_students.text
        if not total_students.isdigit():
            self.show_snackbar("Please enter a valid number of students.")
            return

        self.attendance_data = []
        self.kvs.ids.student_list.clear_widgets()
        for i in range(1, int(total_students) + 1):
            student_layout = BoxLayout(size_hint_y=None, height=40, padding=[10, 0])
            student_label = Label(text=f"Student {i}", size_hint_x=0.6)
            present_button = MDButton(text="Present", on_release=lambda x, i=i: self.mark_attendance(i, "Present"))
            absent_button = MDButton(text="Absent", on_release=lambda x, i=i: self.mark_attendance(i, "Absent"))
            student_layout.add_widget(student_label)
            student_layout.add_widget(present_button)
            student_layout.add_widget(absent_button)
            self.kvs.ids.student_list.add_widget(student_layout)

    def mark_attendance(self, student_number, status):
        for record in self.attendance_data:
            if record["Student Number"] == student_number:
                record["Status"] = status
                self.show_snackbar(f"Student {student_number} marked as {status}")
                return
        self.attendance_data.append({"Student Number": student_number, "Status": status})
        self.show_snackbar(f"Student {student_number} marked as {status}")

    def export_to_excel(self):
        df = pd.DataFrame(self.attendance_data)
        if not os.path.exists("exports"):
            os.makedirs("exports")
        df.to_excel("exports/attendance.xlsx", index=False)
        self.show_snackbar("Attendance data has been exported to attendance.xlsx")

    def show_snackbar(self, message):
        snackbar(text=message).make()

if __name__ == "__main__":
    MainApp().run()

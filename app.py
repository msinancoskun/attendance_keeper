from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import os
import xlrd
import xlsxwriter


class Project:
    global app
    
    def __init__(self, window):
        self.window = window
        self.init_UI()
    
    def init_UI(self):
        # basic configuration
        self.window.title("Attendance Keeper")
        self.window.columnconfigure(0, weight=1)
        self.window.columnconfigure(1, weight=2)
        self.window.columnconfigure(2, weight=1)
        self.window.geometry("640x320")
        
        # Labels
        self.header = Label(self.window, text='Attendance Keeper v1.0', fg="black", font=('Helvetica 18 bold'))
        self.header.grid(row=0, column=0, columnspan=3, sticky='N')

        self.student_label = Label(self.window, text='Select Student Excel File:', font=('Helvetica 12 bold'))
        self.student_label.grid(row=1, column=0)

        self.select_student = Label(self.window, text="Select a Student:", font=('Helvetica 12 bold'))
        self.select_student.grid(row=2, column=0, sticky='W')

        self.section = Label(self.window, text="Section:", font=('Helvetica 12 bold'))
        self.section.grid(row=2, column=1)

        self.box_value = StringVar()
        self.combbox = ttk.Combobox(self.window, textvariable=self.box_value, values=('ENGR 102 01', 'ENGR 102 02', 
                                                                                      'ENGR 102 03', 'ENGR 102 04', 
                                                                                      'ENGR 102 05', 'ENGR 102 06', 
                                                                                      'ENGR 102 07', 'ENGR 102 08', 
                                                                                      'ENGR 102 09', 'ENGR 102 10', 
                                                                                      'ENGR 102 11', 'ENGR 102 12'
                                                                                      ))

        self.combbox.current(0)
        self.combbox.grid(row=3, column=1, sticky='N')

        self.import_button = Button(self.window, text="Import List", font=('Helvetica 12 bold'), command=self.import_file)
        self.import_button.grid(row=1, column=1, ipadx=50)

        self.attended_students = Label(self.window, text="Attended Students:", font=('Helvetica 12 bold'))
        self.attended_students.grid(row=2, column=2, ipadx=50)

        self.student_list = Listbox(self.window, height=5, width=40, selectmode=MULTIPLE)
        self.student_list.grid(row=3, column=0, sticky='W', rowspan=3)

        self.attended_students_list = Listbox(self.window, height=5, width=40)
        self.attended_students_list.grid(row=3, column=2, sticky='E', rowspan=3)

        self.add_button = Button(self.window, text="Add=>", font=('Helvetica 12 bold'), command=self.add_student, width=20)
        self.add_button.grid(row=4, column=1, sticky='N')

        self.remove_button = Button(self.window, text="<=Remove", font=('Helvetica 12 bold'), width=20, command=self.remove_student)
        self.remove_button.grid(row=5, column=1, sticky='N')

        self.file_type = Label(self.window, text="Please select a file type:", font=('Helvetica 9 bold'))
        self.file_type.grid(row=6, column=0, sticky='W')

        self.file_type_value = StringVar()
        self.file_type_combobox = ttk.Combobox(self.window, textvariable=self.file_type_value, width=10)
        self.file_type_combobox['values'] = ('txt', 'xlsx', 'csv')
        self.file_type_combobox.current(0)
        self.file_type_combobox.grid(row=6, column=0, sticky='E')

        self.enter_week = Label(self.window, text="Please enter week:", font=('Helvetica 9 bold'))
        self.enter_week.grid(row=6, column=1)

        self.enter_entry = Entry(self.window)
        self.enter_entry.grid(row=6, column=2, sticky='W')

        self.export_button = Button(self.window, text="Export as File", font=('Helvetica 10 bold'), width=13, command=self.export_file)
        self.export_button.grid(row=6, column=2, sticky='E')

    def import_file(self):
        xlrd.xlsx.ensure_elementtree_imported(False, None)
        xlrd.xlsx.Element_has_iter = True
        self.students = {}

        self.file = filedialog.askopenfile(initialdir=os.getcwd(), filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if self.file:
            file = self.file.name
        path = file
        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)

        for row in range(1, sheet.nrows):
            self.fullname = sheet.cell(row, 1).value
            self.section = sheet.cell(row, 3).value
            self.id = sheet.cell(row, 0).value
            self.students[self.id] = (self.fullname, self.section)

        self.student_list.delete(0, END)
        for id, info in self.students.items():
            self.student_list.insert(END, str(int(id)) + "." + info[0] + ", " + info[1])
            self._filter_section()

    def _filter_section(self):
        selected_list = [i for i in self.student_list.get(0, END)]
        for item in selected_list:
            if item.split(',')[-1].split(' ')[3] != self.combbox.get().split(' ')[2]:
                selected_list.remove(item)

        self.student_list.delete(0, END)
        selected_list.sort()
        for item in selected_list:
            self.student_list.insert(0, item)

    def add_student(self):
        selected_list = [self.student_list.get(i) for i in self.student_list.curselection()]
        for student in selected_list:
            if student not in self.attended_students_list.get(0, END):
                self.attended_students_list.insert(0, student)

    def remove_student(self):
        attended_list = [i for i in self.attended_students_list.get(0, END)]
        selected = self.attended_students_list.curselection()[0]
        for student in attended_list:
            if student == self.attended_students_list.get(selected):
                attended_list.remove(student)
                self.attended_students_list.delete(selected)

    def export_file(self):
        file_type = self.file_type_combobox.get()
        if file_type == 'xlsx':
            self._xls()
        elif file_type == 'txt':
            self._txt()
        else:
            self._csv()

    def _txt(self):
        file_type = self.file_type_combobox.get()
        file_name = self.enter_entry.get()
        section_name = self.combbox.get()
        content = [i for i in self.attended_students_list.get(0, END)]
        path = os.getcwd()
        if not os.path.isfile(os.path.join(path, file_name)):
            file_export = "{} {}.{}".format(section_name, file_name, file_type)
            with open(file_export, 'w', encoding="utf-8") as fp:
                for con in content:
                    fp.write("{}\n".format(str(con)))

    def _xls(self):
        file_type = self.file_type_combobox.get()
        week = self.enter_entry.get()
        section_name = self.combbox.get()
        content = [i for i in self.attended_students_list.get(0, END)]
        path = os.getcwd()
        student_ids = []
        student_names = []
        sections = []
        for item in content:
            student_id = item.split('.')[0]
            student_name = item.split('.')[1].split(', ')[0]
            section_ = item.split(', ')[1]
            student_ids.append(student_id)
            student_names.append(student_name)
            sections.append(section_)

        if not os.path.isfile(os.path.join(path, week)):
            file_export = "{} {}.{}".format(section_name, week, file_type)
            if file_type == 'xlsx':
                workbook = xlsxwriter.Workbook(file_export)
                outSheet = workbook.add_worksheet()
                outSheet.write('A1', 'ID')
                outSheet.write('B1', 'Names')
                outSheet.write('C1', 'Section')
                for item in range(len(student_ids)):
                    outSheet.write(item + 1, 0, student_ids[item])
                    outSheet.write(item + 1, 1, student_names[item])
                    outSheet.write(item + 1, 2, sections[item])
                workbook.close()

    # def _csv(self):
    #     file_type = self.file_type_combobox.get()
    #     file_name = self.enter_entry.get()
    #     section_name = self.combbox.get()
    #     content = [i for i in self.student_list.get(0, END)]
    #     path = os.getcwd()
    #     if not os.path.isfile(os.path.join(path, file_name)):
    #         file_export = "{} {}.{}".format(section_name, file_name, file_type)
    #         with open(file_export, 'w', encoding="utf-8") as fp:
    #             if file_type == 'xls':
    #                 book = Workbook()
    #                 sheet1 = book.add_sheet('Sheet1')
    #                 sheet1.write(0, 0, 'ID')
    #                 sheet1.write(0, 1, 'Name')
    #                 sheet1.write(0, 1, 'Department')


if __name__ == '__main__':
    mainWindow = Tk()
    app = Project(mainWindow)
    mainWindow.mainloop()

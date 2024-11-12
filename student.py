import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import json
import os
from openpyxl import Workbook

class Student:
    def __init__(self, name, age, average_grade):
        self.name = name
        self.age = age
        self.average_grade = average_grade

    def get_info(self):
        return (self.name, self.age, self.average_grade, self.calculate_grade())

    def calculate_grade(self):
        if self.average_grade > 8:
            return "Отлично"
        elif 6 <= self.average_grade <= 8:
            return "Хорошо"
        elif 4 <= self.average_grade < 6:
            return "Удовлетворительно"
        else:
            return "Неудовлетворительно"

class StudentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Management App")

        self.students = []
        self.load_students()

        # Кнопка для выбора всех студентов
        self.select_all_button = tk.Button(root, text="Выбрать все", command=self.select_all_students)
        self.select_all_button.grid(row=0, column=0, padx=10, pady=10)

        # Кнопка для открытия окна добавления студента
        self.add_button = tk.Button(root, text="Добавить студента", command=self.open_add_student_window)
        self.add_button.grid(row=0, column=1, padx=10, pady=10)

        # Кнопка для редактирования выбранного студента
        self.edit_button = tk.Button(root, text="Редактировать выбранное", command=self.edit_selected_student)
        self.edit_button.grid(row=0, column=2, padx=10, pady=10)

        # Кнопка для экспорта в Excel
        self.export_button = tk.Button(root, text="Экспортировать в Excel", command=self.export_to_excel)
        self.export_button.grid(row=0, column=3, padx=10, pady=10)

        # Кнопка для удаления выбранных студентов
        self.delete_button = tk.Button(root, text="Удалить выбранные", command=self.delete_selected_students)
        self.delete_button.grid(row=0, column=4, padx=10, pady=10)

        # Таблица для отображения студентов
        self.table = ttk.Treeview(root, columns=("select", "name", "age", "average_grade", "evaluation"), show="headings")
        self.table.heading("select", text="Выбрать")
        self.table.heading("name", text="Имя")
        self.table.heading("age", text="Возраст")
        self.table.heading("average_grade", text="Средний балл")
        self.table.heading("evaluation", text="Оценка")
        self.table.column("select", width=50)

        # Создание словаря для хранения состояния "чекбоксов"
        self.checkboxes = {}

        self.table.grid(row=1, column=0, columnspan=5, padx=10, pady=10)

        # Отображение данных в таблице после загрузки
        self.populate_table()


    def select_all_students(self):
        # Отмечаем все студенты в таблице
        for item_id in self.table.get_children():
            self.checkboxes[item_id] = True
            self.table.item(item_id, values=("✓", *self.table.item(item_id, "values")[1:]))

    def export_to_excel(self):
        # Проверяем, есть ли выбранные строки
        selected_students = [
            self.table.item(item_id, "values")[1:]  # Пропускаем чекбокс
            for item_id, checked in self.checkboxes.items() if checked
        ]

        if not selected_students:
            messagebox.showwarning("Ошибка", "Выберите строки для экспорта.")
            return

        # Создаем новый Excel файл
        wb = Workbook()
        ws = wb.active
        ws.title = "Студенты"

        # Заголовки для Excel файла
        headers = ["Имя", "Возраст", "Средний балл", "Оценка"]
        ws.append(headers)

        # Заполняем данные в Excel
        for student_data in selected_students:
            ws.append(student_data)

        # Сохраняем файл Excel
        file_path = os.path.join(os.path.dirname(__file__), "students_export.xlsx")
        wb.save(file_path)

        # Сообщение об успешном экспорте
        messagebox.showinfo("Экспорт завершен", f"Данные успешно экспортированы в файл {file_path}")

    def open_add_student_window(self, student=None, item_id=None):
        # Создание нового окна для добавления или редактирования студента
        self.add_window = tk.Toplevel(self.root)
        self.add_window.title("Добавить студента" if student is None else "Редактировать студента")

        # Поля для ввода данных студента
        tk.Label(self.add_window, text="Имя:").grid(row=0, column=0, padx=5, pady=5)
        self.name_entry = tk.Entry(self.add_window)
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(self.add_window, text="Возраст:").grid(row=1, column=0, padx=5, pady=5)
        self.age_entry = tk.Entry(self.add_window)
        self.age_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(self.add_window, text="Средний балл:").grid(row=2, column=0, padx=5, pady=5)
        self.grade_entry = tk.Entry(self.add_window)
        self.grade_entry.grid(row=2, column=1, padx=5, pady=5)

        # Если студент передан для редактирования, заполняем поля текущими значениями
        if student:
            self.name_entry.insert(0, student.name)
            self.age_entry.insert(0, student.age)
            self.grade_entry.insert(0, student.average_grade)

        # Кнопка для сохранения данных
        save_button = tk.Button(self.add_window, text="Сохранить", command=lambda: self.save_student(item_id))
        save_button.grid(row=3, column=0, columnspan=2, pady=10)

    def save_student(self, item_id=None):
        # Получение данных из полей ввода
        name = self.name_entry.get()
        age = self.age_entry.get()
        average_grade = self.grade_entry.get()

        # Проверка на пустые поля
        if not name or not age or not average_grade:
            messagebox.showwarning("Ошибка", "Все поля должны быть заполнены!")
            return

        try:
            age = int(age)
            average_grade = float(average_grade)
        except ValueError:
            messagebox.showwarning("Ошибка", "Возраст должен быть числом, а средний балл - числом с точкой!")
            return

        if item_id is not None:
            # Удаляем старую запись из таблицы и списка студентов
            values = self.table.item(item_id, "values")[1:]  # Убираем чекбокс
            self.students = [s for s in self.students if (s.name.strip() != values[0].strip() or str(s.age) != str(values[1]) or str(s.average_grade) != str(values[2]))]
            self.table.delete(item_id)
            del self.checkboxes[item_id]

        # Добавление студента как новой записи
        student = Student(name, age, average_grade)
        self.students.append(student)
        new_item_id = self.table.insert("", "end", values=("", *student.get_info()))
        self.checkboxes[new_item_id] = False

        # Сохранение данных в JSON-файл
        self.save_students_to_file()

        # Закрытие окна добавления/редактирования
        self.add_window.destroy()


    def save_students_to_file(self):
        # Подготовка данных для записи в JSON
        data = [
            {
                "name": student.name,
                "age": student.age,
                "average_grade": student.average_grade
            }
            for student in self.students
        ]

        # Отладочный вывод данных перед записью
        print("Данные, которые будут записаны в JSON:", data)  # Отладочный вывод

        # Определение пути к файлу
        file_path = os.path.join(os.path.dirname(__file__), "students.json")
        
        # Сохранение в JSON файл
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)


    def load_students(self):
        # Определение пути к файлу
        file_path = os.path.join(os.path.dirname(__file__), "students.json")
        
        # Загрузка данных из JSON-файла, если файл существует
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                for item in data:
                    student = Student(item["name"], item["age"], item["average_grade"])
                    self.students.append(student)

    def populate_table(self):
        # Заполнение таблицы данными, загруженными из JSON-файла
        for idx, student in enumerate(self.students):
            item_id = self.table.insert("", "end", values=("", *student.get_info()))
            self.checkboxes[item_id] = False

        # Подключение события для изменения состояния "чекбоксов"
        self.table.bind("<Button-1>", self.toggle_checkbox)

    def toggle_checkbox(self, event):
        # Определение элемента по координатам клика
        row_id = self.table.identify_row(event.y)
        column_id = self.table.identify_column(event.x)

        # Проверка, что клик был в столбце "select"
        if column_id == "#1" and row_id in self.checkboxes:
            # Переключение состояния чекбокса
            self.checkboxes[row_id] = not self.checkboxes[row_id]
            new_value = "✓" if self.checkboxes[row_id] else ""
            self.table.item(row_id, values=(new_value, *self.table.item(row_id, "values")[1:]))

    def delete_selected_students(self):
        # Получаем все элементы, отмеченные для удаления
        items_to_delete = [item_id for item_id, checked in self.checkboxes.items() if checked]

        # Проверяем, есть ли что удалять
        if not items_to_delete:
            messagebox.showwarning("Ошибка", "Выберите строки для удаления.")
            return

        # Удаление отмеченных записей
        for item_id in items_to_delete:
            # Получаем данные студента для удаления из списка `self.students`
            values = self.table.item(item_id, "values")[1:]  # Убираем чекбокс

            # Удаление студента из списка `self.students`
            self.students = [
                s for s in self.students
                if not (s.name.strip() == values[0].strip() and str(s.age) == str(values[1]) and str(s.average_grade) == str(values[2]))
            ]

            # Удаление элемента из таблицы и словаря чекбоксов
            self.table.delete(item_id)
            del self.checkboxes[item_id]

        # Проверка результата после удаления
        print("Оставшиеся студенты после удаления:", [s.get_info() for s in self.students])  # Отладочный вывод

        # Сохранение обновленного списка студентов в JSON-файл
        self.save_students_to_file()

        # Сообщение об успешном удалении
        messagebox.showinfo("Удаление", "Выбранные студенты были удалены.")

    def edit_selected_student(self):
        # Проверка на выбранные строки
        selected_items = [item_id for item_id, checked in self.checkboxes.items() if checked]
        if len(selected_items) != 1:
            messagebox.showwarning("Ошибка", "Выберите одну строку для редактирования.")
            return

        # Получаем item_id выбранного элемента
        item_id = selected_items[0]
        
        # Получаем данные выбранного студента из таблицы
        values = self.table.item(item_id, "values")[1:]  # Убираем чекбокс

        # Поиск студента в списке self.students
        student_index = None
        for i, s in enumerate(self.students):
            if s.name.strip() == values[0].strip() and str(s.age) == str(values[1]) and str(s.average_grade) == str(values[2]):
                student_index = i
                break

        # Если студент не найден, показываем ошибку
        if student_index is None:
            messagebox.showerror("Ошибка", "Не удалось найти выбранного студента.")
            return

        # Открытие окна для редактирования с текущими данными студента, передаем item_id
        self.open_add_student_window(self.students[student_index], item_id)


if __name__ == "__main__":
    root = tk.Tk()
    app = StudentApp(root)
    root.mainloop()

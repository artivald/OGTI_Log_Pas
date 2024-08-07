import sqlite3
import string
import random
from openpyxl import Workbook
from PyQt5.QtWidgets import QApplication, QMainWindow, QMenu, QMenuBar, QAction, QFileDialog, QTableWidget, \
    QTableWidgetItem, QVBoxLayout, QWidget, QHeaderView, QDialog, QLabel, QLineEdit, QPushButton, QComboBox, QHBoxLayout
import sys
from transliterate import translit
import csv

DB_PATH = None

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_db_path = None
        self.initUI()

    def initUI(self):
        # Создаем меню
        menubar = self.menuBar()
        file_menu = menubar.addMenu('Файл')
        edit_menu = menubar.addMenu('Редактирование')

        # Действия для меню "Файл"
        select_db_action = QAction('Выбрать базу данных', self)
        select_db_action.triggered.connect(self.select_db)
        file_menu.addAction(select_db_action)

        # Действия для меню "Редактирование"
        delete_action = QAction('Удалить', self)
        delete_action.triggered.connect(self.delete_user)
        add_action = QAction('Добавить', self)
        add_action.triggered.connect(self.add_user_form)
        edit_action = QAction('Редактировать', self)
        edit_action.triggered.connect(self.edit_user_dialog)
        edit_menu.addAction(delete_action)
        edit_menu.addAction(add_action)
        edit_menu.addAction(edit_action)

        # Действия для меню "Загрузка/Выгрузка"
        load_export_menu = QMenu('Загрузка/Выгрузка', self)
        export_to_excel_action = QAction('Выгрузить в Excel', self)
        export_to_excel_action.triggered.connect(self.export_selected_to_excel)
        load_export_menu.addAction(export_to_excel_action)
        export_to_CSV_action = QAction('Выгрузить в CSV', self)
        export_to_CSV_action.triggered.connect(self.export_selected_to_CSV)
        load_export_menu.addAction(export_to_CSV_action)
        menubar.addMenu(load_export_menu)

        self.setWindowTitle('Управление пользователями')
        self.setGeometry(100, 100, 920, 600)

    def select_db(self):
        global DB_PATH
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл базы данных", "",
                                                   "Файлы баз данных (*.db);;Все файлы (*.*)")
        if file_path:
            self.current_db_path = file_path
            DB_PATH = file_path
            self.load_data(file_path)

    def load_data(self, file_path):
        conn = sqlite3.connect(file_path)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM user")
        data = cursor.fetchall()

        table_area = QWidget()
        layout = QVBoxLayout(table_area)

        filter_layout = QHBoxLayout()
        self.last_name_filter = QLineEdit()
        self.first_name_filter = QLineEdit()
        self.middle_name_filter = QLineEdit()
        self.subdivision_filter = QComboBox()
        self.position_filter = QComboBox()
        self.faculty_filter = QComboBox()

        subdivisions = self.get_divisions()
        positions = self.get_posts()
        faculties = self.get_faculties()
        self.subdivision_filter.addItem("Все")
        self.subdivision_filter.addItems(subdivisions)
        self.position_filter.addItem("Все")
        self.position_filter.addItems(positions)
        self.faculty_filter.addItem("Все")
        self.faculty_filter.addItems(faculties)

        filter_layout.addWidget(QLabel("Фамилия:"))
        filter_layout.addWidget(self.last_name_filter)
        filter_layout.addWidget(QLabel("Имя:"))
        filter_layout.addWidget(self.first_name_filter)
        filter_layout.addWidget(QLabel("Отчество:"))
        filter_layout.addWidget(self.middle_name_filter)
        filter_layout.addWidget(QLabel("Подразделение:"))
        filter_layout.addWidget(self.subdivision_filter)
        filter_layout.addWidget(QLabel("Должность:"))
        filter_layout.addWidget(self.position_filter)
        filter_layout.addWidget(QLabel("Факультет:"))
        filter_layout.addWidget(self.faculty_filter)

        filter_button = QPushButton("Применить фильтры")
        filter_button.clicked.connect(self.apply_filters)
        filter_layout.addWidget(filter_button)

        layout.addLayout(filter_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(len(data[0]))
        self.table.setRowCount(len(data))
        self.table.setHorizontalHeaderLabels(
            ["ID", "Фамилия", "Имя", "Отчество", "Логин", "Пароль", "Подразделение", "Должность", "Факультет"])

        for row, row_data in enumerate(data):
            for col, item in enumerate(row_data):
                self.table.setItem(row, col, QTableWidgetItem(str(item)))

        self.table.hideColumn(0)  # Скрываем столбец с ID

        self.table.setSelectionBehavior(QTableWidget.SelectRows)  # Выделение строк

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.sectionClicked.connect(self.sort_table)

        layout.addWidget(self.table)
        table_area.setLayout(layout)
        self.setCentralWidget(table_area)
        conn.close()

    def apply_filters(self):
        conn = sqlite3.connect(self.current_db_path)
        cursor = conn.cursor()
        query = "SELECT * FROM user WHERE 1=1"
        params = []

        if self.last_name_filter.text():
            query += " AND surname LIKE ?"
            params.append(f"%{self.last_name_filter.text()}%")
        if self.first_name_filter.text():
            query += " AND name LIKE ?"
            params.append(f"%{self.first_name_filter.text()}%")
        if self.middle_name_filter.text():
            query += " AND patronymic LIKE ?"
            params.append(f"%{self.middle_name_filter.text()}%")
        if self.subdivision_filter.currentText() != "Все":
            query += " AND division = ?"
            params.append(self.subdivision_filter.currentText())
        if self.position_filter.currentText() != "Все":
            query += " AND post = ?"
            params.append(self.position_filter.currentText())
        if self.faculty_filter.currentText() != "Все":
            query += " AND faculty = ?"
            params.append(self.faculty_filter.currentText())

        cursor.execute(query, params)
        data = cursor.fetchall()

        self.table.setRowCount(len(data))
        for row, row_data in enumerate(data):
            for col, item in enumerate(row_data):
                self.table.setItem(row, col, QTableWidgetItem(str(item)))

        self.table.hideColumn(0)
        conn.close()

    def sort_table(self, column):
        sort_order = self.table.horizontalHeader().sortIndicatorOrder()
        self.table.sortItems(column, sort_order)

    def refresh_table(self):
        pass

    def delete_user(self):
        selected_rows = self.table.selectedItems()
        if not selected_rows:
            return
        ids_to_delete = set()
        for item in selected_rows:
            row = item.row()
            id_item = self.table.item(row, 0)
            if id_item:
                ids_to_delete.add(int(id_item.text()))

        conn = sqlite3.connect(self.current_db_path)
        cursor = conn.cursor()
        for user_id in ids_to_delete:
            cursor.execute("DELETE FROM user WHERE id=?", (user_id,))
        conn.commit()
        conn.close()
        self.load_data(self.current_db_path)

    def add_user_form(self):
        add_dialog = QDialog(self)
        add_dialog.setWindowTitle("Добавить пользователя")

        last_name_label = QLabel("Фамилия:")
        last_name_edit = QLineEdit()
        first_name_label = QLabel("Имя:")
        first_name_edit = QLineEdit()
        middle_name_label = QLabel("Отчество:")
        middle_name_edit = QLineEdit()
        login_label = QLabel("Логин:")
        login_edit = QLineEdit()
        password_label = QLabel("Пароль:")
        password_edit = QLineEdit()
        subdivision_label = QLabel("Подразделение:")
        subdivision_edit = QComboBox()
        position_label = QLabel("Должность:")
        position_edit = QComboBox()
        faculty_label = QLabel("Факультет:")
        faculty_edit = QComboBox()

        add_button = QPushButton("Добавить")
        cancel_button = QPushButton("Отмена")
        generate_login_button = QPushButton("Сгенерировать логин")
        generate_password_button = QPushButton("Сгенерировать пароль")

        subdivisions = self.get_divisions()
        positions = self.get_posts()
        faculties = self.get_faculties()
        subdivision_edit.addItems(subdivisions)
        position_edit.addItems(positions)
        faculty_edit.addItems(faculties)

        layout = QVBoxLayout()
        layout.addWidget(last_name_label)
        layout.addWidget(last_name_edit)
        layout.addWidget(first_name_label)
        layout.addWidget(first_name_edit)
        layout.addWidget(middle_name_label)
        layout.addWidget(middle_name_edit)
        layout.addWidget(login_label)
        layout.addWidget(login_edit)
        layout.addWidget(generate_login_button)
        layout.addWidget(password_label)
        layout.addWidget(password_edit)
        layout.addWidget(generate_password_button)
        layout.addWidget(subdivision_label)
        layout.addWidget(subdivision_edit)
        layout.addWidget(position_label)
        layout.addWidget(position_edit)
        layout.addWidget(faculty_label)
        layout.addWidget(faculty_edit)
        layout.addWidget(add_button)
        layout.addWidget(cancel_button)

        add_dialog.setLayout(layout)

        generate_login_button.clicked.connect(lambda: self.generate_login(last_name_edit, first_name_edit, middle_name_edit, login_edit))
        generate_password_button.clicked.connect(lambda: self.generate_password(password_edit))

        add_button.clicked.connect(lambda: self.add_user(
            last_name_edit.text(),
            first_name_edit.text(),
            middle_name_edit.text(),
            login_edit.text(),
            password_edit.text(),
            subdivision_edit.currentText(),
            position_edit.currentText(),
            faculty_edit.currentText(),
            add_dialog
        ))

        cancel_button.clicked.connect(add_dialog.close)

        add_dialog.exec_()

    def generate_login(self, last_name_edit, first_name_edit, middle_name_edit, login_edit):
        last_name = last_name_edit.text().strip()
        first_name = first_name_edit.text().strip()
        middle_name = middle_name_edit.text().strip()
        login = translit(f"{last_name}.{first_name[0]}.{middle_name[0]}", 'ru', reversed=True).lower()
        login_edit.setText(login)

    def generate_password(self, password_edit):
        alphabet = string.ascii_letters + string.digits
        password = ''.join(random.choice(alphabet) for _ in range(8))
        password_edit.setText(password)

    def add_user(self, last_name, first_name, middle_name, login, password, subdivision, position, faculty, dialog):
        conn = sqlite3.connect(self.current_db_path)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO user (surname, name, patronymic, login, password, division, post, faculty)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (last_name, first_name, middle_name, login, password, subdivision, position, faculty))
        conn.commit()
        conn.close()
        self.load_data(self.current_db_path)
        dialog.close()

    def edit_user_dialog(self):
        selected_rows = self.table.selectedItems()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        id_item = self.table.item(row, 0)

        edit_dialog = QDialog(self)
        edit_dialog.setWindowTitle("Редактировать пользователя")

        last_name_label = QLabel("Фамилия:")
        last_name_edit = QLineEdit(self.table.item(row, 1).text())
        first_name_label = QLabel("Имя:")
        first_name_edit = QLineEdit(self.table.item(row, 2).text())
        middle_name_label = QLabel("Отчество:")
        middle_name_edit = QLineEdit(self.table.item(row, 3).text())
        login_label = QLabel("Логин:")
        login_edit = QLineEdit(self.table.item(row, 4).text())
        password_label = QLabel("Пароль:")
        password_edit = QLineEdit(self.table.item(row, 5).text())
        subdivision_label = QLabel("Подразделение:")
        subdivision_edit = QComboBox()
        position_label = QLabel("Должность:")
        position_edit = QComboBox()
        faculty_label = QLabel("Факультет:")
        faculty_edit = QComboBox()

        edit_button = QPushButton("Сохранить")
        cancel_button = QPushButton("Отмена")

        subdivisions = self.get_divisions()
        positions = self.get_posts()
        faculties = self.get_faculties()
        subdivision_edit.addItems(subdivisions)
        subdivision_edit.setCurrentText(self.table.item(row, 6).text())
        position_edit.addItems(positions)
        position_edit.setCurrentText(self.table.item(row, 7).text())
        faculty_edit.addItems(faculties)
        faculty_edit.setCurrentText(self.table.item(row, 8).text())

        layout = QVBoxLayout()
        layout.addWidget(last_name_label)
        layout.addWidget(last_name_edit)
        layout.addWidget(first_name_label)
        layout.addWidget(first_name_edit)
        layout.addWidget(middle_name_label)
        layout.addWidget(middle_name_edit)
        layout.addWidget(login_label)
        layout.addWidget(login_edit)
        layout.addWidget(password_label)
        layout.addWidget(password_edit)
        layout.addWidget(subdivision_label)
        layout.addWidget(subdivision_edit)
        layout.addWidget(position_label)
        layout.addWidget(position_edit)
        layout.addWidget(faculty_label)
        layout.addWidget(faculty_edit)
        layout.addWidget(edit_button)
        layout.addWidget(cancel_button)

        edit_dialog.setLayout(layout)

        edit_button.clicked.connect(lambda: self.edit_user(
            id_item.text(),
            last_name_edit.text(),
            first_name_edit.text(),
            middle_name_edit.text(),
            login_edit.text(),
            password_edit.text(),
            subdivision_edit.currentText(),
            position_edit.currentText(),
            faculty_edit.currentText(),
            edit_dialog
        ))

        cancel_button.clicked.connect(edit_dialog.close)

        edit_dialog.exec_()

    def edit_user(self, user_id, last_name, first_name, middle_name, login, password, subdivision, position, faculty, dialog):
        conn = sqlite3.connect(self.current_db_path)
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE user
            SET surname = ?, name = ?, patronymic = ?, login = ?, password = ?, division = ?, post = ?, faculty = ?
            WHERE id = ?
        """, (last_name, first_name, middle_name, login, password, subdivision, position, faculty, user_id))
        conn.commit()
        conn.close()
        self.load_data(self.current_db_path)
        dialog.close()

    def get_divisions(self):
        conn = sqlite3.connect(self.current_db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT division FROM user")
        divisions = [row[0] for row in cursor.fetchall()]
        conn.close()
        return divisions

    def get_posts(self):
        conn = sqlite3.connect(self.current_db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT post FROM user")
        posts = [row[0] for row in cursor.fetchall()]
        conn.close()
        return posts

    def get_faculties(self):
        conn = sqlite3.connect(self.current_db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT faculty FROM user")
        faculties = [row[0] for row in cursor.fetchall()]
        conn.close()
        return faculties

    def export_selected_to_excel(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как", "", "Excel Files (*.xlsx);;All Files (*)",
                                                   options=options)

        # Добавление расширения, если его нет
        if file_path and not file_path.endswith('.xlsx'):
            file_path += '.xlsx'

        if not file_path:
            return

        selected_rows = set()
        for item in selected_items:
            selected_rows.add(item.row())

        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["ID", "Фамилия", "Имя", "Отчество", "Логин", "Пароль", "Подразделение", "Должность", "Факультет"])

        for row in selected_rows:
            row_data = [self.table.item(row, col).text() for col in range(self.table.columnCount())]
            sheet.append(row_data)

        workbook.save(file_path)

    def export_selected_to_CSV(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить как", "", "CSV Files (*.csv);;All Files (*)")

        # Добавление расширения, если его нет
        if file_path and not file_path.endswith('.csv'):
            file_path += '.csv'

        if not file_path:
            return

        selected_rows = set()
        for item in selected_items:
            selected_rows.add(item.row())

        with open(file_path, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(
                ["ID", "Фамилия", "Имя", "Отчество", "Логин", "Пароль", "Подразделение", "Должность", "Факультет"])

            for row in selected_rows:
                row_data = [self.table.item(row, col).text() for col in range(self.table.columnCount())]
                writer.writerow(row_data)


def main():
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()

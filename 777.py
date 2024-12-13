import sys
import sqlite3
import pandas as pd  # Импортируем библиотеку pandas для работы с Excel
from PyQt5.QtWidgets import QFileDialog  # Импортируем QFileDialog для выбора файла
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QStackedWidget, QWidget, QVBoxLayout,
    QLabel, QLineEdit, QPushButton, QMessageBox, QListWidget,
    QHBoxLayout, QCheckBox, QComboBox
)

# Путь к файлу для хранения данных пользователя
AUTH_FILE = "auth.txt"

# Розовый стиль для большинства окон
pink_style = """
QMainWindow, QWidget {
    background-color: #ffc0cb;  /* Розовый фон */
}

QLabel {
    font-size: 16px;
    font-weight: bold;
    color: #333;
    margin: 5px;
}

QLineEdit {
    border: 1px solid #ff69b4; 
    border-radius: 5px;
    padding: 8px;
    background-color: #fff;  
    font-weight: bold;
}

QPushButton {
    background-color: #ff69b4;  
    color: white;
    border: none;
    border-radius: 15px;  
    padding: 8px; 
    font-size: 16px;
    font-weight: bold;          
}

QPushButton:hover {
    background-color: #db7093;  
}

QListWidget {
    background-color: #fff;
    border: 1px solid #ff69b4;
    border-radius: 5px;
    font-weight: bold;          
}

QComboBox {
    border: 1px solid #ff69b4;
    border-radius: 5px;
    padding: 5px;
    font-weight: bold;          
}

QCheckBox {
    margin: 5px;
    font-weight: bold;          
}
"""

# Стиль только для окна пользователя
customer_style = """
QWidget {
    background-color: #ffc0cb;  /* Розовый фон для CustomerWindow */
}
"""


def initialize_database():
    conn = sqlite3.connect("tours.db")
    cursor = conn.cursor()

    # Создание таблиц
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Tours (
            Tours_id INTEGER PRIMARY KEY,
            Name TEXT,
            Price INTEGER NOT NULL,
            Tours_Date TEXT NOT NULL,
            Duration TEXT NOT NULL,
            City TEXT NOT NULL,
            Country_id INTEGER NOT NULL,
            Hotel_id INTEGER NOT NULL,
            Flights_id INTEGER NOT NULL,
            Client_id INTEGER NOT NULL,
            Description TEXT  -- добавлен столбец Description
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Sales (
            Sales_id INTEGER PRIMARY KEY,
            Date_of_agreement TEXT,
            Quantity INTEGER,
            client_id INTEGER,
            tours_id INTEGER,
            phone TEXT,
            payment_status TEXT DEFAULT 'неоплачено'  -- добавлен столбец для статуса оплаты
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Refunds (
            Refunds_id INTEGER PRIMARY KEY AUTOINCREMENT,
            Sales_id INTEGER NOT NULL,
            Client_id INTEGER NOT NULL,
            Tours_id INTEGER NOT NULL,
            Date_of_agreement TEXT NOT NULL,
            Quantity INTEGER NOT NULL,
            Reason TEXT NOT NULL,
            FOREIGN KEY (Sales_id) REFERENCES Sales(Sales_id),
            FOREIGN KEY (Client_id) REFERENCES users(id),
            FOREIGN KEY (Tours_id) REFERENCES Tours(Tours_id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Country (
            Country_id INTEGER PRIMARY KEY,
            Name TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Hotel (
            Hotel_id INTEGER PRIMARY KEY,
            Name TEXT NOT NULL,
            Country_id INTEGER NOT NULL,
            FOREIGN KEY (Country_id) REFERENCES Country(Country_id) ON DELETE CASCADE ON UPDATE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Flights (
            Flights_id INTEGER PRIMARY KEY,
            Departure_date TEXT NOT NULL,
            Departure_time TEXT NOT NULL,
            Flight_time INTEGER NOT NULL,
            Baggage TEXT NOT NULL,
            Country_id INTEGER NOT NULL
        )
    """)

    # Заполнение таблицы стран
    cursor.execute("SELECT COUNT(*) FROM Country")
    count = cursor.fetchone()[0]
    if count == 0:  # Проверяем, если таблица стран пуста
        countries = [('Турция',), ('Бразилия',), ('Египет',), ('Италия',), ('Испания',), ('Франция',)]
        cursor.executemany("INSERT INTO Country (Name) VALUES (?)", countries)

    # Заполнение таблицы отелей
    cursor.execute("SELECT COUNT(*) FROM Hotel")
    count = cursor.fetchone()[0]
    if count == 0:  # Проверяем, если таблица отелей пуста
        hotels = [
            ('Гранд Отель', 1),  # Турция
            ('Солнечный Берег', 1),
            ('Морской Ветер', 1),
            ('Копакабана', 2),  # Бразилия
            ('Рио Отель', 2),
            ('Прайя Отель', 2),
            ('Кайро Палас', 3),  # Египет
            ('Пирамид Гранд', 3),
            ('Сфинкс Отель', 3),
            ('Венецианка', 4),  # Италия
            ('Римский Двор', 4),
            ('Тоскана Отель', 4),
            ('Барселона Плаза', 5),  # Испания
            ('Мадрид Палаццо', 5),
            ('Севильский Гранд', 5),
            ('Люкс Отель', 6),  # Франция
            ('Шеридан', 6),
            ('Парижская Долина', 6)
        ]
        cursor.executemany("INSERT INTO Hotel (Name, Country_id) VALUES (?, ?)", hotels)

    # Заполнение таблицы рейсов
    cursor.execute("SELECT COUNT(*) FROM Flights")
    count = cursor.fetchone()[0]
    if count == 0:  # Проверяем, если таблица рейсов пуста
        flights = [
            ('2023-05-01', '10:00', 3, '20kg', 1),  # Вылет в Турцию
            ('2023-06-15', '15:00', 5, '25kg', 1),
            ('2023-07-10', '12:30', 4, '30kg', 1),
            ('2023-05-15', '08:00', 4, '20kg', 2),  # Вылет в Бразилию
            ('2023-07-05', '13:00', 3, '15kg', 2),
            ('2023-08-20', '16:00', 2, '25kg', 2),
            ('2023-04-20', '09:30', 4, '10kg', 3),  # Вылет в Египет
            ('2023-05-30', '14:45', 6, '20kg', 3),
            ('2023-06-18', '11:15', 5, '30kg', 3),
            ('2023-05-01', '07:30', 2, '15kg', 4),  # Вылет в Италию
            ('2023-06-10', '09:00', 2, '20kg', 4),
            ('2023-07-07', '12:00', 3, '25kg', 4),
            ('2023-05-20', '11:30', 5, '30kg', 5),  # Вылет в Испанию
            ('2023-07-15', '15:30', 3, '20kg', 5),
            ('2023-08-10', '10:15', 2, '25kg', 5),
            ('2023-05-01', '18:00', 1, '20kg', 6),  # Вылет во Францию
            ('2023-06-15', '16:30', 3, '30kg', 6),
            ('2023-07-04', '09:00', 1, '15kg', 6)
        ]
        cursor.executemany(
            "INSERT INTO Flights (Departure_date, Departure_time, Flight_time, Baggage, Country_id) VALUES (?, ?, ?, ?, ?)",
            flights)

    create_default_admin(cursor)  # Проверяем и создаем админа по умолчанию

    conn.commit()
    conn.close()


def create_default_admin(cursor):
    # Проверяем, есть ли администратор
    cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
    count = cursor.fetchone()[0]
    print(f"Количество администраторов: {count}")  # Для отладки

    if count == 0:
        # Добавляем администратора по умолчанию
        admin_name = "Admin"
        admin_email = "admin@example.com"
        admin_password = "admin123"  # Задайте безопасный пароль
        cursor.execute("INSERT INTO users (name, email, password, role) VALUES (?, ?, ?, 'admin')",
                       (admin_name, admin_email, admin_password))
        print("Администратор по умолчанию добавлен.")
    else:
        print("Администратор уже существует.")


# Функция для записи почты и пароля в файл
def save_auth_data(email, password):
    with open(AUTH_FILE, "w") as f:
        f.write(f"{email}\n{password}")


# Функция для чтения почты и пароля из файла
def load_auth_data():
    try:
        with open(AUTH_FILE, "r") as f:
            email = f.readline().strip()
            password = f.readline().strip()
            return email, password
    except FileNotFoundError:
        return None, None


class LoginWindow(QWidget):
    def __init__(self, switch_window):
        super().__init__()
        self.switch_window = switch_window
        self.setWindowTitle("Вход")
        self.setStyleSheet(pink_style)

        layout = QVBoxLayout()
        self.email_input = QLineEdit(self)
        self.email_input.setPlaceholderText("Введите ваш email")
        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText("Введите ваш пароль")
        self.password_input.setEchoMode(QLineEdit.Password)

        # Загружаем последние сохранённые данные
        saved_email, saved_password = load_auth_data()
        if saved_email and saved_password:
            self.email_input.setText(saved_email)
            self.password_input.setText(saved_password)

        login_button = QPushButton("Войти")
        login_button.clicked.connect(self.login)

        register_button = QPushButton("Зарегистрироваться")
        register_button.clicked.connect(self.register)

        layout.addWidget(QLabel("Вход"))
        layout.addWidget(self.email_input)
        layout.addWidget(self.password_input)
        layout.addWidget(login_button)
        layout.addWidget(register_button)

        self.setLayout(layout)

    def login(self):
        email = self.email_input.text().strip()
        password = self.password_input.text().strip()

        if not email or not password:
            QMessageBox.warning(self, "Ошибка", "Все поля должны быть заполнены!")
            return

        conn = sqlite3.connect("tours.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE email = ? AND password = ?", (email, password))
        user = cursor.fetchone()
        conn.close()

        if user:
            QMessageBox.information(self, "Успешно", f"Добро пожаловать, {user[1]}!")
            save_auth_data(email, password)  # Сохраним данные
            self.switch_window("admin" if user[4] == "admin" else "customer", user[0])
        else:
            QMessageBox.warning(self, "Ошибка", "Неверный email или пароль!")

    def register(self):
        self.switch_window("register")


class RegisterWindow(QWidget):
    def __init__(self, switch_window):
        super().__init__()
        self.switch_window = switch_window
        self.setWindowTitle("Регистрация")
        self.setStyleSheet(pink_style)

        layout = QVBoxLayout()

        self.name_input = QLineEdit(self)
        self.name_input.setPlaceholderText("Введите ваше имя")
        self.email_input = QLineEdit(self)
        self.email_input.setPlaceholderText("Введите ваш email")
        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText("Введите ваш пароль")
        self.password_input.setEchoMode(QLineEdit.Password)

        register_button = QPushButton("Зарегистрироваться")
        register_button.clicked.connect(self.register)
        back_button = QPushButton("Назад")
        back_button.clicked.connect(self.back_to_login)

        layout.addWidget(QLabel("Регистрация"))
        layout.addWidget(self.name_input)
        layout.addWidget(self.email_input)
        layout.addWidget(self.password_input)
        layout.addWidget(register_button)
        layout.addWidget(back_button)

        self.setLayout(layout)

    def register(self):
        name = self.name_input.text().strip()
        email = self.email_input.text().strip()
        password = self.password_input.text().strip()

        if not name or not email or not password:
            QMessageBox.warning(self, "Ошибка", "Все поля должны быть заполнены!")
            return

        conn = sqlite3.connect("tours.db")
        cursor = conn.cursor()
        try:
            cursor.execute("INSERT INTO users (name, email, password, role) VALUES (?, ?, ?, 'customer')",
                           (name, email, password))
            conn.commit()
            QMessageBox.information(self, "Успешно", "Регистрация прошла успешно!")
            self.switch_window("login")
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Ошибка", "Email уже существует!")
        finally:
            conn.close()

    def back_to_login(self):
        self.switch_window("login")


class BookingWindow(QWidget):
    def __init__(self, tour_id, user_id):
        super().__init__()
        self.tour_id = tour_id
        self.user_id = user_id
        self.setWindowTitle("Оформление тура")

        layout = QVBoxLayout()
        self.name_input = QLineEdit(self)
        self.name_input.setPlaceholderText("ФИО")
        self.email_input = QLineEdit(self)
        self.email_input.setPlaceholderText("Email")
        self.passport_input = QLineEdit(self)
        self.passport_input.setPlaceholderText("Паспортные данные")
        self.phone_input = QLineEdit(self)
        self.phone_input.setPlaceholderText("Номер телефона")

        self.quantity_input = QLineEdit(self)
        self.quantity_input.setPlaceholderText("Количество человек")
        self.payment_method = QComboBox(self)
        self.payment_method.addItems(["Карта", "Наличные", "Банковский перевод"])

        book_button = QPushButton("Оформить тур")
        book_button.clicked.connect(self.book_tour)

        layout.addWidget(QLabel("ФИО:"))
        layout.addWidget(self.name_input)
        layout.addWidget(QLabel("Email:"))
        layout.addWidget(self.email_input)
        layout.addWidget(QLabel("Паспортные данные:"))
        layout.addWidget(self.passport_input)
        layout.addWidget(QLabel("Номер телефона:"))
        layout.addWidget(self.phone_input)
        layout.addWidget(QLabel("Количество человек:"))
        layout.addWidget(self.quantity_input)
        layout.addWidget(QLabel("Способ оплаты:"))
        layout.addWidget(self.payment_method)
        layout.addWidget(book_button)

        self.setLayout(layout)

    def book_tour(self):
        # Подтверждение оформления
        reply = QMessageBox.question(self, 'Подтверждение оформления',
                                     "Вы уверены, что хотите оформить тур?",
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)

        if reply == QMessageBox.No:
            return  # Прерываем оформление, если пользователь нажал "Нет"

        # Если пользователь подтверждает оформление
        name = self.name_input.text().strip()
        email = self.email_input.text().strip()
        passport = self.passport_input.text().strip()
        phone = self.phone_input.text().strip()
        quantity = self.quantity_input.text().strip()

        if not name or not email or not passport or not phone or not quantity:
            QMessageBox.warning(self, "Ошибка", "Все поля должны быть заполнены!")
            return

        conn = sqlite3.connect("tours.db")
        cursor = conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO Sales (Date_of_agreement, Quantity, client_id, tours_id, phone, payment_status) VALUES (datetime('now'), ?, ?, ?, ?, 'неоплачено')",
                (quantity, self.user_id, self.tour_id, phone))
            conn.commit()

            QMessageBox.information(self, "Успешно",
                                    f"Тур успешно оформлен!\n\nВ ближайшее время с Вами свяжется турагент.\nНомер телефона: {phone}\n\nЕсли в течение 30 минут Вам не позвонят, просьба позвонить по этому номеру телефона +234 325.")
            self.close()
        except sqlite3.Error as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось оформить тур: {str(e)}")
        finally:
            conn.close()


class TourDetailWindow(QWidget):
    def __init__(self, tour_id):
        super().__init__()
        self.tour_id = tour_id
        self.setWindowTitle("Детали тура")
        self.setStyleSheet(pink_style)

        layout = QVBoxLayout()

        self.tour_info = QLabel(self)
        layout.addWidget(self.tour_info)

        back_button = QPushButton("Назад")
        back_button.clicked.connect(self.close)
        layout.addWidget(back_button)

        self.setLayout(layout)
        self.load_tour_details()

    def load_tour_details(self):
        conn = sqlite3.connect("tours.db")
        cursor = conn.cursor()

        cursor.execute("""
            SELECT Tours.Name, Tours.Price, Tours.Description, Tours.Duration, Tours.City, Hotel.Name
            FROM Tours
            JOIN Hotel ON Tours.Hotel_id = Hotel.Hotel_id
            WHERE Tours.Tours_id = ?
        """, (self.tour_id,))

        tour = cursor.fetchone()
        conn.close()

        if tour:
            price_formatted = f"{tour[1]:,.2f} ₽"  # Форматируем цену с двумя десятичными знаками
            self.tour_info.setText(
                f"Тур: {tour[0]}\n"
                f"Цена: {price_formatted}\n"
                f"Описание: {tour[2]}\n"
                f"Продолжительность: {tour[3]}\n"
                f"Город: {tour[4]}\n"
                f"Отель: {tour[5]}"
            )
        else:
            self.tour_info.setText("Информация о туре не найдена.")


class MyToursWindow(QWidget):
    def __init__(self, user_id):
        super().__init__()
        self.user_id = user_id

        layout = QVBoxLayout()
        self.tour_list = QListWidget(self)
        layout.addWidget(QLabel("Забронированные туры:"))
        layout.addWidget(self.tour_list)

        refund_button = QPushButton("Запросить возврат тура")
        refund_button.clicked.connect(self.request_refund)
        layout.addWidget(refund_button)

        update_button = QPushButton("Обновить список туров")
        update_button.clicked.connect(self.load_my_tours)
        layout.addWidget(update_button)

        self.setLayout(layout)
        self.load_my_tours()

    def load_my_tours(self):
        self.tour_list.clear()
        # Получение данных из БД
        conn = sqlite3.connect("tours.db")
        cursor = conn.cursor()
        cursor.execute("""
            SELECT Sales.Sales_id, Tours.Name, Price, Sales.phone, Sales.payment_status FROM Sales 
            JOIN Tours ON Sales.tours_id = Tours.Tours_id 
            WHERE Sales.client_id = ?
        """, (self.user_id,))
        my_tours = cursor.fetchall()
        conn.close()

        for tour in my_tours:
            price_formatted = f"{tour[2]:,.2f} ₽"  # Форматируем цену с двумя десятичными знаками
            self.tour_list.addItem(
                f"ID: {tour[0]} | Тур: {tour[1]} | Цена: {price_formatted} | Номер телефона: {tour[3]} | Статус: {tour[4]}")

    def request_refund(self):
        selected_item = self.tour_list.currentItem()
        if selected_item:
            QMessageBox.information(self, "Запрос возврата", "Запрос на возврат успешно отправлен!")
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите тур для запроса возврата!")

class ReportsWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Отчеты об оплаченных турах")
        self.setStyleSheet(pink_style)

        layout = QVBoxLayout()

        self.report_list = QListWidget(self)
        layout.addWidget(QLabel("Отчеты об оплаченных турах:"))
        layout.addWidget(self.report_list)

        self.setLayout(layout)

        self.load_reports()  # Загружаем отчеты сразу при создании окна

    def load_reports(self):
        self.report_list.clear()
        try:
            conn = sqlite3.connect("tours.db")
            cursor = conn.cursor()
            cursor.execute("""
                SELECT 
                    Tours.Tours_id, 
                    Users.name, 
                    Users.email, 
                    Sales.Date_of_agreement
                FROM Sales 
                JOIN Tours ON Sales.tours_id = Tours.Tours_id 
                JOIN Users ON Sales.client_id = Users.id
                WHERE Sales.payment_status = 'оплачено'
            """)
            reports = cursor.fetchall()

            if reports:
                for report in reports:
                    self.report_list.addItem(
                        f"ID тура: {report[0]} | ФИО: {report[1]} | Email: {report[2]} | Дата договора: {report[3]}"
                    )
            else:
                self.report_list.addItem("Нет оплаченных туров.")

        except sqlite3.Error as e:
            print(f"Database error in load_reports: {e}")
        finally:
            if conn:
                conn.close()

class CustomerWindow(QWidget):
    def __init__(self, user_id, switch_window):
        super().__init__()
        self.user_id = user_id
        self.switch_window = switch_window
        self.setWindowTitle("Пользовательская панель")
        self.setStyleSheet(customer_style)
        layout = QVBoxLayout()

        # Чекбоксы для фильтрации
        self.country_checkboxes = {}
        layout.addWidget(QLabel("Выберите страны для фильтрации:"))

        conn = sqlite3.connect("tours.db")
        cursor = conn.cursor()
        cursor.execute("SELECT Country_id, Name FROM Country")
        countries = cursor.fetchall()
        conn.close()

        for country in countries:
            checkbox = QCheckBox(country[1])
            checkbox.stateChanged.connect(self.filter_tours)
            self.country_checkboxes[country[0]] = checkbox
            layout.addWidget(checkbox)

        # Список туров
        self.tour_list = QListWidget(self)
        self.tour_list.setFixedHeight(200)  # Уменьшение высоты списка туров
        self.tour_list.itemDoubleClicked.connect(self.open_tour_detail)
        layout.addWidget(QLabel("Доступные туры"))
        layout.addWidget(self.tour_list)

        # Кнопка "Оформить тур"
        self.tour_button = QPushButton("Оформить тур")
        self.tour_button.clicked.connect(self.open_booking_window)
        layout.addWidget(self.tour_button)

        # Кнопка "Обновить туры"
        refresh_button = QPushButton("Обновить туры")
        refresh_button.clicked.connect(self.load_tours)
        layout.addWidget(refresh_button)

        # Кнопка "Выйти"
        logout_button = QPushButton("Выйти")
        logout_button.clicked.connect(self.logout)
        layout.addWidget(logout_button)

        self.setLayout(layout)
        self.load_tours()

    def open_booking_window(self):
        # Открытие окна для оформления тура
        selected_items = self.tour_list.selectedItems()
        if selected_items:
            selected_tour_id = selected_items[0].text().split("|")[0].split(":")[1].strip()
            self.booking_window = BookingWindow(tour_id=selected_tour_id, user_id=self.user_id)
            self.booking_window.show()
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите тур для оформления!")

    def load_tours(self):
        # Загрузка всех доступных туров
        self.tour_list.clear()
        conn = sqlite3.connect("tours.db")
        cursor = conn.cursor()
        cursor.execute("""
            SELECT Tours_id, Name, Price, Description FROM Tours
        """)
        tours = cursor.fetchall()
        conn.close()

        for tour in tours:
            price_formatted = f"{tour[2]:,.2f} ₽"  # Форматируем цену с двумя десятичными знаками
            self.tour_list.addItem(f"ID: {tour[0]} | Тур: {tour[1]} | Цена: {price_formatted} | Описание: {tour[3]}")

    def filter_tours(self):
        # Фильтрация туров по странам
        selected_country_ids = [country_id for country_id, checkbox in self.country_checkboxes.items() if
                                checkbox.isChecked()]

        if selected_country_ids:
            conn = sqlite3.connect("tours.db")
            cursor = conn.cursor()
            cursor.execute("""
                SELECT Tours_id, Name, Price, Description FROM Tours WHERE Country_id IN ({seq})
            """.format(seq=','.join(['?'] * len(selected_country_ids))), selected_country_ids)
            tours = cursor.fetchall()
            conn.close()
        else:
            self.load_tours()
            return

        self.tour_list.clear()
        for tour in tours:
            self.tour_list.addItem(f"ID: {tour[0]} | Тур: {tour[1]} | Цена: {tour[2]} | Описание: {tour[3]}")

    def open_tour_detail(self, item):
        # Открытие окна деталей тура
        selected_tour_id = item.text().split("|")[0].split(":")[1].strip()
        self.tour_detail_window = TourDetailWindow(tour_id=selected_tour_id)
        self.tour_detail_window.show()

    def logout(self):
        # Выход из учетной записи
        self.switch_window("login")


class AdminWindow(QWidget):
    def __init__(self, switch_window):
        super().__init__()
        self.switch_window = switch_window
        self.setWindowTitle("Панель администратора")
        self.setStyleSheet(pink_style)

        layout = QVBoxLayout()

        self.user_list = QListWidget(self)
        layout.addWidget(QLabel("Список пользователей"))
        layout.addWidget(self.user_list)

        button_layout = QHBoxLayout()
        refresh_button = QPushButton("Возвраты")
        refresh_button.clicked.connect(self.show_refunds_window)
        logout_button = QPushButton("Выйти")
        logout_button.clicked.connect(self.logout)

        button_layout.addWidget(refresh_button)
        button_layout.addWidget(logout_button)

        self.booking_list = QListWidget(self)
        layout.addWidget(QLabel("Забронированные туры"))
        layout.addWidget(self.booking_list)

        booking_button_layout = QHBoxLayout()
        refresh_bookings_button = QPushButton("Обновить бронирования")
        refresh_bookings_button.clicked.connect(lambda: self.load_bookings())
        delete_booking_button = QPushButton("Удалить бронирование")
        delete_booking_button.clicked.connect(self.delete_booking)

        change_payment_button = QPushButton("Изменить статус оплаты")
        change_payment_button.clicked.connect(self.change_payment_status)

        booking_button_layout.addWidget(refresh_bookings_button)
        booking_button_layout.addWidget(delete_booking_button)
        booking_button_layout.addWidget(change_payment_button)

        layout.addLayout(booking_button_layout)
        layout.addLayout(button_layout)
        self.setLayout(layout)

        self.load_users()
        self.load_bookings()
        # Кнопка "Отчеты"
        reports_button = QPushButton("Отчеты")
        reports_button.clicked.connect(self.show_reports_window)  # Подключаем открытие окна отчетов
        layout.addWidget(reports_button)

        self.setLayout(layout)

        self.load_users()
        self.load_bookings()

    def load_users(self):
        self.user_list.clear()
        conn = sqlite3.connect("tours.db")
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, email, role FROM users")
        users = cursor.fetchall()
        conn.close()

        for user in users:
            self.user_list.addItem(f"ID: {user[0]} | Имя: {user[1]} | Email: {user[2]} | Роль: {user[3]}")

    def show_reports_window(self):
        # Логика для открытия окна отчетов
        self.reports_window = ReportsWindow()  # Предполагается, что ReportsWindow уже реализован
        self.reports_window.show()

    def load_bookings(self):
        self.booking_list.clear()
        conn = sqlite3.connect("tours.db")
        cursor = conn.cursor()
        cursor.execute("""
            SELECT Sales.Sales_id, Tours.Name, Users.name, Sales.phone, Sales.payment_status 
            FROM Sales 
            JOIN Tours ON Sales.tours_id = Tours.Tours_id 
            JOIN users ON Sales.client_id = Users.id
        """)
        bookings = cursor.fetchall()
        conn.close()

        if not bookings:
            print("Нет забронированных туров.")
        else:
            for booking in bookings:
                self.booking_list.addItem(
                    f"ID: {booking[0]} | Тур: {booking[1]} | Клиент: {booking[2]} | Номер телефона: {booking[3]} | Статус оплаты: {booking[4]}")

        print(f"Загружено бронирований: {len(bookings)}")

    def change_payment_status(self):
        selected_item = self.booking_list.currentItem()
        if selected_item:
            booking_id = selected_item.text().split("|")[0].split(":")[1].strip()
            conn = sqlite3.connect("tours.db")
            cursor = conn.cursor()
            cursor.execute("SELECT payment_status FROM Sales WHERE Sales_id = ?", (booking_id,))
            current_status = cursor.fetchone()
            current_status_text = current_status[0] if current_status else None

            new_status = "оплачено" if current_status_text == "неоплачено" else "неоплачено"

            reply = QMessageBox.question(self, 'Изменение статуса оплаты',
                                         f"Вы уверены, что хотите изменить статус оплаты с '{current_status_text}' на '{new_status}'?",
                                         QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.No)

            if reply == QMessageBox.Yes:
                cursor.execute("UPDATE Sales SET payment_status = ? WHERE Sales_id = ?", (new_status, booking_id))
                conn.commit()
                QMessageBox.information(self, "Успешно", f"Статус оплаты изменен на '{new_status}'.")
                self.load_bookings()  # Обновляем список бронирований
            conn.close()
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите запись для изменения статуса!")

    def delete_booking(self):
        selected_item = self.booking_list.currentItem()
        if selected_item:
            booking_id = selected_item.text().split("|")[0].split(":")[1].strip()
            conn = sqlite3.connect("tours.db")
            cursor = conn.cursor()
            cursor.execute("DELETE FROM Sales WHERE Sales_id = ?", (booking_id,))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Успешно", "Бронирование удалено!")
            self.load_bookings()  # Обновляем список бронирований
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, выберите запись для удаления!")

    def logout(self):
        self.switch_window("login")

    def show_refunds_window(self):
        self.refunds_window = RefundsWindow()
        self.refunds_window.show()

class ReportsWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Отчеты об оплаченных турах")
        self.setStyleSheet(pink_style)

        layout = QVBoxLayout()

        self.report_list = QListWidget(self)
        layout.addWidget(QLabel("Отчеты об оплаченных турах:"))
        layout.addWidget(self.report_list)

        # Кнопка для экспорта отчетов в Excel
        export_button = QPushButton("Экспортировать в Excel")
        export_button.clicked.connect(self.export_to_excel)  # Подключаем метод экспорта
        layout.addWidget(export_button)

        self.setLayout(layout)

        self.load_reports()  # Загружаем отчеты сразу при создании окна

    def load_reports(self):
        self.report_list.clear()
        try:
            conn = sqlite3.connect("tours.db")
            cursor = conn.cursor()
            cursor.execute(""" 
                SELECT 
                    Sales.Sales_id, 
                    Tours.Name, 
                    Users.name, 
                    Sales.Date_of_agreement 
                FROM Sales 
                JOIN Tours ON Sales.tours_id = Tours.Tours_id 
                JOIN Users ON Sales.client_id = Users.id 
                WHERE Sales.payment_status = 'оплачено' 
            """)
            reports = cursor.fetchall()

            if reports:
                for report in reports:
                    self.report_list.addItem(
                        f"ID продажи: {report[0]} | Тур: {report[1]} | Клиент: {report[2]} | Дата договора: {report[3]}"
                    )
            else:
                self.report_list.addItem("Нет оплаченных туров.")

        except sqlite3.Error as e:
            print(f"Database error in load_reports: {e}")
        finally:
            if conn:
                conn.close()

    def export_to_excel(self):
        # Экспортируем отчеты в файл Excel
        try:
            conn = sqlite3.connect("tours.db")
            cursor = conn.cursor()
            cursor.execute(""" 
                SELECT 
                    Sales.Sales_id, 
                    Tours.Name, 
                    Users.name, 
                    Sales.Date_of_agreement 
                FROM Sales 
                JOIN Tours ON Sales.tours_id = Tours.Tours_id 
                JOIN Users ON Sales.client_id = Users.id 
                WHERE Sales.payment_status = 'оплачено' 
            """)
            reports = cursor.fetchall()

            if reports:
                # Создаем DataFrame из полученных данных
                df = pd.DataFrame(reports, columns=["ID продажи", "Тур", "Клиент", "Дата договора"])

                # Открываем диалог для выбора места сохранения файла
                file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", "", "Excel Files (*.xlsx)")
                if file_path:
                    df.to_excel(file_path, index=False)  # Экспортируем в Excel
                    QMessageBox.information(self, "Успех", "Отчет успешно экспортирован в Excel!")
            else:
                QMessageBox.warning(self, "Предупреждение", "Нет оплаченных туров для экспорта.")

        except Exception as e:
            print(f"Error exporting to Excel: {e}")
            QMessageBox.critical(self, "Ошибка", "Произошла ошибка при экспорте отчета.")
        finally:
            if conn:
                conn.close()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Главная")
        self.setFixedSize(800, 600)

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.login_window = LoginWindow(self.switch_window)
        self.register_window = RegisterWindow(self.switch_window)
        self.admin_window = AdminWindow(self.switch_window)
        self.customer_window = CustomerWindow(user_id=None, switch_window=self.switch_window)

        self.stack.addWidget(self.login_window)
        self.stack.addWidget(self.register_window)
        self.stack.addWidget(self.admin_window)
        self.stack.addWidget(self.customer_window)

    def switch_window(self, window_name, user_id=None):
        if window_name == "login":
            self.stack.setCurrentWidget(self.login_window)
        elif window_name == "register":
            self.stack.setCurrentWidget(self.register_window)
        elif window_name == "admin":
            self.stack.setCurrentWidget(self.admin_window)
        elif window_name == "customer":
            self.customer_window.user_id = user_id  # Обновляем user_id
            self.stack.setCurrentWidget(self.customer_window)

if __name__ == "__main__":
    initialize_database()
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    initialize_database()
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
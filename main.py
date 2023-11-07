# Импорт необходимых библиотек
import os  # Используется для взаимодействия с операционной системой
import numpy as np  # Популярная библиотека для численных расчетов
import shutil  # Используется для работы с файлами и папками
import pandas as pd  # Библиотека для анализа и обработки данных
from selenium import webdriver  # Инструмент для автоматизации действий в веб-браузере
from selenium.webdriver.chrome.service import Service  # Сервис для управления ChromeDriver
from selenium.webdriver.common.by import By  # Используется для указания методов поиска элементов на веб-страницах
from selenium.webdriver.common.action_chains import ActionChains  # Инструмент для автоматизации комплексных действий пользователя
import time  # Библиотека для работы со временем
import re  # Модуль для работы с регулярными выражениями
from datetime import datetime  # Используется для работы с датами и временем
import subprocess  # Используется для запуска новых процессов
import logging  # Библиотека для ведения логов

# Конфигурация логгирования
logging.basicConfig(
    level=logging.INFO,  # Устанавливаем уровень логгирования - INFO
    format='%(asctime)s - %(levelname)s - %(message)s',  # Формат вывода логов
    handlers=[
        logging.FileHandler('changes.log', mode='w'),  # Логи будут записываться в файл changes.log
        logging.StreamHandler()  # Логи также будут выводиться в стандартный поток вывода
    ]
)

# Функция для извлечения данных о компаниях с веб-страницы
def extract_company_data(url):
    # Настройка опций для Chrome
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--disable-gpu')  # Отключаем GPU для ускорения загрузки
    chrome_options.add_argument('--headless')  # Режим без графического интерфейса для серверного использования

    # Инициализация драйвера Chrome с заданными опциями
    driver_service = Service('/usr/bin/chromedriver') #Указать путь к chromedariver
    browser = webdriver.Chrome(service=driver_service, options=chrome_options)

    try:
        # Открытие указанной веб-страницы
        browser.get(url)

        # Ожидание для полной загрузки страницы
        time.sleep(10)  # Ждем 10 секунд, а не 20, как указано в комментарии

        # Поиск элемента input по его атрибуту ng-model на странице
        input_element = browser.find_element(By.CSS_SELECTOR, 'input[ng-model="vm.iprFilter.esName"]')

        # Создание объекта для выполнения серии действий с элементом
        actions = ActionChains(browser)

        # Перемещение к элементу и клик по нему
        actions.move_to_element(input_element).click().perform()

        # Ожидание для загрузки контента после действия
        time.sleep(5)

        # Получение данных о компаниях с веб-страницы
        company_elements = browser.find_elements(By.CSS_SELECTOR, 'li')  # Необходимо уточнить CSS-селектор

        companies = []  # Список для хранения данных о компаниях
        for elem in company_elements:
            text = elem.text
            match = re.match(r"^(.+) \((\d+)\)$", text)  # Проверка текста элемента на соответствие шаблону
            if match:
                # Извлечение названия компании и ИНН из текста
                company_name = match.group(1)
                inn = match.group(2)  # ИНН сохраняется как строка

        # Произведение замен в названии компании для укорочения формы записи
        company_name = company_name.replace("Общество с ограниченной ответственностью", "ООО")
        company_name = company_name.replace("Закрытое акционерное", "ЗАО")
        company_name = company_name.replace("Акционерное общество", "АО")
        company_name = company_name.replace("Муниципальное унитарное предприятие", "МУП")
        company_name = company_name.replace("Акционерное Общество", "АО")
        company_name = company_name.replace("Общество с ограниченной ответственность", "ООО")

        # Добавление статуса и текущей даты/времени к данным компании
        status = "Действующая"  # Статус компании, возможно, потребуется получать из другого источника
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')  # Форматирование текущей даты и времени

        # Добавление собранной информации о компании в список
        companies.append({"NAME_GOSUSLUGI": company_name, "INN": inn, "STATUS": status, "DATE/TIME": timestamp})

    except Exception as e:
        # В случае ошибки выводится сообщение
        print(f"An error occurred: {e}")
    finally:
        # Закрытие браузера после завершения работы
        browser.quit()

    # Возврат списка с данными о компаниях
    return companies

# Функция для чтения файла Excel с данными о компаниях
def read_excel(file_name="companies.xlsx"):
    # Проверка наличия файла
    if os.path.exists(file_name):
        # Чтение файла с указанием типа данных для колонки ИНН
        return pd.read_excel(file_name, dtype={'INN': str}, engine='openpyxl')
    else:
        # В случае отсутствия файла возвращается пустой DataFrame
        return pd.DataFrame()

# Функция для создания резервной копии старой версии файла
def backup_old_version(file_name="companies.xlsx"):
    # Проверка существования папки backup, если нет - создание
    if not os.path.exists("backup"):
        os.mkdir("backup")
    # Форматирование текущей даты и времени для создания уникального имени файла
    timestamp = datetime.now().strftime('%d-%m-%Y %H-%M-%S')
    # Копирование файла в папку backup с новым именем
    shutil.copy(file_name, os.path.join("backup", f"{timestamp}.xlsx"))

# Функция для сравнения и обновления данных о компаниях
def compare_and_update(companies):
    # Список ИНН, которые необходимо исключить из обработки
    exclude_inns = ["999999999", "52561056945", "3023011567", "5012060636", "007704726225"]

    # Чтение данных из файла Excel
    old_data = read_excel()

    if not old_data.empty:
        # Создание резервной копии файла, если он не пуст
        backup_old_version()

    # Удаление метки 'NEW' у компаний, которые уже есть в списке
    common_inns = old_data[old_data["INN"].isin(companies["INN"])]["INN"]
    old_data.loc[old_data["INN"].isin(common_inns), 'COLOR'] = np.nan

    # Выделение новых компаний, исключая заданные ИНН
    added = companies[~companies["INN"].isin(old_data["INN"]) & ~companies["INN"].isin(exclude_inns)].copy()

    # Присвоение номеров новым компаниям
    if not old_data.empty:
        max_number = old_data['№'].fillna(0).astype(int).max()
    else:
        max_number = 0

    added['№'] = range(max_number + 1, max_number + 1 + len(added))

    # Пометка новых компаний
    added['COLOR'] = 'NEW'

    # Объединение новых и старых данных
    combined_data = pd.concat([old_data, added], ignore_index=True)

    # Сохранение обновленного файла
    save_to_excel(combined_data, "companies.xlsx")

# Функция для сохранения данных в файл Excel с учетом цветовой маркировки
def save_to_excel(companies, file_name):
    # Создание ExcelWriter для записи данных
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    companies.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']

    # Форматирование для выделения новых данных цветом
    green_format = writer.book.add_format({'bg_color': '#00FF00'})

    # Применение форматирования к новым данным
    for index, row in companies.iterrows():
        color = row.get('COLOR')
        if color == 'NEW':
            for col_num, value in enumerate(row):
                # Пропуск столбца 'COLOR' и пустых ячеек
                if pd.notnull(value):
                    worksheet.write(index + 1, col_num, value, green_format)

    # Закрытие и сохранение файла
    writer.close()

# Основная часть скрипта
url = "https://invest.gosuslugi.ru/epgu-forum/#/ipr"  # URL для извлечения данных
companies = extract_company_data(url)  # Получение данных о компаниях
companies = pd.DataFrame(companies)  # Преобразование списка в DataFrame
compare_and_update(companies)  # Сравнение и обновление данных

# Запуск дополнительных скриптов после основной логики
subprocess.run(["python", "CHECK_STATUS.py"])  # Скрипт для проверки статуса
subprocess.run(["python", "YA_DISK.py"])  # Скрипт для работы с Яндекс.Диском
subprocess.run(["python", "GOOGLE_SHEETS.py"])  # Скрипт для работы с Google Sheets

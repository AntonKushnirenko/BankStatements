from kivymd.app import MDApp
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.screen import MDScreen
import gspread
google_sheet_name = "Выписки"
service_account_filename = "service_account.json"
worksheet_name = "Лист1"  # Название листа (страницы), которое выбирается снизу таблицы.
starting_directory = "/home/anton"
ooo_search_words = ["ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО"]
ip_search_words = ["ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ", "ИП"]

class MainScreen(MDScreen):
    def __init__(self, **kwargs):
        super(MainScreen, self).__init__(**kwargs)
        self.manager_open = False
        self.file_manager = MDFileManager(exit_manager=self.exit_manager, select_path=self.select_path)
        self.lines = []  # Строки текстового файла

    # Выбор файла
    def open_file_manager(self, *args):
        self.file_manager.show(starting_directory)  # Здесь указываем начальную директорию

    # Этот метод будет вызываться при выборе файла
    def select_path(self, path):
        print("Путь: ", path)
        # Проверяем текстовый ли файл
        if self.is_txt(path):
            # Читаем содержимое
            with open(path, "r", encoding='Windows-1251') as file:
                self.lines = file.read().splitlines()
                #print(f"Текст: {self.lines}")
                #print(self.get_required_values());

                self.exit_manager()

    # Проверяем текстовый ли файл
    def is_txt(self, path):
        if path.endswith(('.txt', '.doc', '.docx', '.odt')):
            return True
        else:
            print("Неподдерживаемый формат файла")
            return False

    def exit_manager(self, *args):
        self.manager_open = False
        self.file_manager.close()

    # Получаем наш расчетный счет, на который идет приток денег
    def get_income_checking_account(self):
        for line in self.lines:
            if line.startswith('РасчСчет='):
                income_checking_account = line.replace('РасчСчет=', "")
                break
        return income_checking_account

    # Получаем значения необходимые для вычисления данных, которые потом будем вносить в таблицу
    def get_required_values(self):
        dates = []  # Дата
        beneficiary_banks = []  # ПолучательБанк1
        payer_banks = []  # ПлательщикБанк1
        recipients = []  # Получатель1
        payers = []  # Плательщик1
        sums_of_money = []  # Сумма
        recipients_accounts = []  # ПолучательСчет
        payers_accounts = []  # ПлательщикСчет
        purposes_of_payments = []  # НазначениеПлатежа

        is_document_section_started = False
        for line in self.lines:
            if line.startswith('СекцияДокумент='):
                is_document_section_started = True

            if is_document_section_started:
                if line.startswith('Дата='):
                    date = line.replace('Дата=', "")
                    dates.append(date)

                if line.startswith('ПолучательБанк1='):
                    beneficiary_bank = line.replace('ПолучательБанк1=', "")
                    beneficiary_banks.append(beneficiary_bank)

                if line.startswith('ПлательщикБанк1='):
                    payer_bank = line.replace('ПлательщикБанк1=', "")
                    payer_banks.append(payer_bank)

                if line.startswith(('Получатель=', 'Получатель1=')):
                    if line.startswith('Получатель='):
                        recipient = line.replace('Получатель=', "")
                    if line.startswith('Получатель1='):
                        recipient = line.replace('Получатель1=', "")
                    recipients.append(recipient)

                if line.startswith(('Плательщик=', 'Плательщик1=')):
                    if line.startswith('Плательщик='):
                        payer = line.replace('Плательщик=', "")
                    if line.startswith('Плательщик1='):
                        payer = line.replace('Плательщик1=', "")
                    payers.append(payer)

                if line.startswith('Сумма='):
                    amount_of_money = line.replace('Сумма=', "")
                    sums_of_money.append(amount_of_money)

                if line.startswith('ПолучательСчет='):
                    recipient_account = line.replace('ПолучательСчет=', "")
                    recipients_accounts.append(recipient_account)

                if line.startswith('ПлательщикСчет='):
                    payer_account = line.replace('ПлательщикСчет=', "")
                    payers_accounts.append(payer_account)

                if line.startswith('НазначениеПлатежа='):
                    purpose_of_payment = line.replace('НазначениеПлатежа=', "")
                    purposes_of_payments.append(purpose_of_payment)

        required_values = list(map(list, zip(dates, beneficiary_banks, payer_banks, recipients,
                                             payers, sums_of_money, recipients_accounts,
                                             payers_accounts, purposes_of_payments)))
        return required_values

    # Формируем данные которые будем уже непосредственно вносить в таблицу
    def get_data_to_upload(self, required_values, income_checking_account):
        data_to_upload = []
        for values in required_values:
            if values[6] == income_checking_account:
                is_income = True  # Приток
            elif values[7] == income_checking_account:
                is_income = False  # Отток

            payment_date = values[0]  # Дата оплаты
            comment = values[8]  # Комментарий

            if is_income:
                payment_type = values[1]  # Тип оплаты (ПолучательБанк1)

                if any(search_word in values[3] for search_word in ooo_search_words):
                # search_word in string это функция, которая выполняется для всех search_word в search_words
                # any возращает true, если хоть одно значение true
                    legal_entity = "ООО"  # Юр лицо (Получатель1)
                elif any(search_word in values[3] for search_word in ip_search_words):
                    legal_entity = "ИП"  # Юр лицо (Получатель1)

                income = values[5] # Приток
                outcome = ''  # Отток
                counterparty = values[4]  # Контрагент (Плательщик1)

            if not is_income:
                payment_type = values[2]  # Тип оплаты (ПлательщикБанк1)

                if any(search_word in values[4] for search_word in ooo_search_words):
                    legal_entity = "ООО"  # Юр лицо (Плательщик1)
                elif any(search_word in values[4] for search_word in ip_search_words):
                    legal_entity = "ИП"  # Юр лицо (Плательщик1)

                income = ''  # Приток
                outcome = values[5]  # Отток
                counterparty = values[3]  # Контрагент (Получатель1)

            data_to_upload.append([payment_date, payment_type, legal_entity, income, outcome, counterparty, comment])
        return data_to_upload

    # Поиск следующей свободной строки в таблице
    def next_available_row(self, worksheet):
        str_list = list(filter(None, worksheet.col_values(1))) # Берем все значения из первого столбика, потом ...?
        return len(str_list) + 1

    # Работа с гугл таблицами
    def upload_to_googledrive(self):
        service_account = gspread.service_account(filename=service_account_filename)
        spreadsheet = service_account.open(google_sheet_name)
        worksheet = spreadsheet.worksheet(worksheet_name)

        required_values = self.get_required_values()
        income_checking_account = self.get_income_checking_account()  # РасчСчет
        data_to_upload = self.get_data_to_upload(required_values, income_checking_account)
        if data_to_upload:  # Проверяем не пустой ли файл
            worksheet.update(f"A{self.next_available_row(worksheet)}:G{self.next_available_row(worksheet) + len(data_to_upload)}", data_to_upload)

        else:
            print("Файл не выбран или содержимое отсутствует.")

class BankStatementsApp(MDApp):
    def build(self):
        return MainScreen()
BankStatementsApp().run()


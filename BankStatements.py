from kivymd.app import MDApp
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.screen import MDScreen
import gspread
google_sheet_name = "Выписки"
service_account_filename = "service_account.json"
worksheet_name = "Лист1"  # Название листа (страницы), которое выбирается снизу таблицы.
starting_directory = "/home/anton"

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

    # Работа с гугл таблицами
    def upload_to_googledrive(self):
        service_account = gspread.service_account(filename=service_account_filename)
        spreadsheet = service_account.open(google_sheet_name)
        worksheet = spreadsheet.worksheet(worksheet_name)
        if self.lines != []:  # Проверяем не пустой ли файл
            worksheet.update(f"A{self.next_available_row(worksheet)}",self.lines)

        else:
            print("Файл не выбран или содержимое отсутствует.")

    # Поиск следующей свободной строки в таблице
    def next_available_row(self, worksheet):
        str_list = list(filter(None, worksheet.col_values(1))) # Берем все значения из первого столбика, потом ...?
        return len(str_list) + 1

    """
    def upload_dates_to_googledrive(self, dates):
        service_account = gspread.service_account(filename=service_account_filename)
        spreadsheet = service_account.open(google_sheet_name)
        worksheet = spreadsheet.worksheet(worksheet_name)
        if dates != []:
            print(f"A{self.next_available_row(worksheet)}:A{self.next_available_row(worksheet) + len(dates)}")
            print(dates)
            worksheet.update(f"A{self.next_available_row(worksheet)}:A{self.next_available_row(worksheet)+len(dates)}", dates)

        else:
            print("Файл не выбран или содержимое отсутствует.")
    """

    # Получаем наш расчетный счет, на который идет приток денег
    def get_income_checking_account(self):
        for line in self.lines:
            if line.startswith('РасчСчет='):
                income_checking_account = line.replace('РасчСчет=', "")
                break
        return income_checking_account

    # Получаем значения необходимые для вычисления данных, которые потом будем вносить в таблицу
    def get_required_values(self):
        income_checking_account = self.get_income_checking_account(self.lines)  # РасчСчет
        dates = []  # Дата
        payer_banks = []  # ПлательщикБанк1
        beneficiary_banks = []  # ПолучательБанк1
        recipients = []  # Получатель1
        payers = []  # Плательщик1
        sums_of_money = []  # Сумма
        recipients_accounts = []  # ПолучательСчет
        payers_accounts = []  # ПлательщикСчет
        purposes_of_payments = []  # НазначениеПлатежа

        for line in self.lines:
            if line.startswith('Дата='):
                date = line.replace('Дата=', "")
                dates.append(date)

            if line.startswith('ПлательщикБанк1='):
                payer_bank = line.replace('ПлательщикБанк1=', "")
                payer_banks.append(payer_bank)

            if line.startswith('ПолучательБанк1='):
                beneficiary_bank = line.replace('ПолучательБанк1=', "")
                beneficiary_banks.append(beneficiary_bank)

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

        required_values = list(map(list, zip(dates, payer_banks, beneficiary_banks, recipients,
                                             payers, amount_of_money, recipients_accounts,
                                             payers_accounts, purposes_of_payments)))
        return required_values, income_checking_account

    # Формируем данные которые будем уже непосредственно вносить в таблицу
    def get_data_to_upload(self):
        pass

class BankStatementsApp(MDApp):
    def build(self):
        return MainScreen()
BankStatementsApp().run()


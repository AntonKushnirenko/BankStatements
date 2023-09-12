from kivymd.app import MDApp
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.screen import MDScreen
from kivymd.uix.label import MDLabel
from kivymd.uix.snackbar import Snackbar
import gspread

google_sheet_name = "Выписки"
service_account_filename = "service_account.json"
worksheet_name = "Лист1"  # Название листа (страницы), которое выбирается снизу таблицы.
starting_directory = "/home/anton"
ooo_search_words = ["ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО"]
ip_search_words = ["ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ", "ИП"]
rate_search_words = ["курс", "Курс сделки", "курс ЦБ"]

class MainScreen(MDScreen):
    def __init__(self, **kwargs):
        super(MainScreen, self).__init__(**kwargs)
        self.manager_open = False
        self.file_manager = MDFileManager(exit_manager=self.exit_manager, select_path=self.select_path)
        self.lines = []  # Строки текстового файла
        self.selected_file_label = None
        self.not_txt_snackbar = None
        self.data_error_snackbar = None

    # Выбор файла
    def open_file_manager(self):
        self.file_manager.show(starting_directory)  # Здесь указываем начальную директорию

    # Этот метод будет вызываться при выборе файла
    def select_path(self, path):
        print("Путь: ", path)
        app = BankStatementsApp.get_running_app()
        # Проверяем текстовый ли файл
        if self.is_txt(path):
            # Читаем содержимое
            with open(path, "r", encoding='Windows-1251') as file:
                self.lines = file.read().splitlines()
                if not self.selected_file_label:
                    self.selected_file_label = MDLabel(text=f"Выбран файл: {path}", halign='center',
                                                       pos_hint={"center_x": 0.5, "center_y": 0.45},
                                                       font_style="H5")
                    self.add_widget(self.selected_file_label)
                elif self.selected_file_label:
                    self.selected_file_label.text = f"Выбран файл: {path}"
                self.exit_manager(self)
        else:
            # Вылезающее уведомление с ошибкой формата файла снизу
            if not self.not_txt_snackbar:
                self.not_txt_snackbar = Snackbar(text="Неподдерживаемый формат файла!",
                                                 font_size=app.font_size_value,
                                                 duration=1, size_hint_x=0.8,
                                                 pos_hint={"center_x": 0.5, "center_y": 0.1})
            self.not_txt_snackbar.open()

    # Проверяем текстовый ли файл
    @staticmethod
    def is_txt(path):
        # if path.endswith(('.txt', '.doc', '.docx', '.odt')):
        if path.endswith('.txt'):
            return True
        else:
            print("Неподдерживаемый формат файла")
            return False

    def exit_manager(self, instance):
        self.manager_open = False
        self.file_manager.close()

    # Получаем наш расчетный счет, на который идет приток денег
    def get_income_checking_account(self):
        for line in self.lines:
            if line.startswith('РасчСчет='):
                income_checking_account = line.replace('РасчСчет=', "")
                return income_checking_account
        print("Ошибка в получении расчетного счета ")
        return ""

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
                    elif line.startswith('Получатель1='):
                        recipient = line.replace('Получатель1=', "")
                    else:
                        recipient = ""
                        print("Ошибка в Получателе")
                    recipients.append(recipient)

                if line.startswith(('Плательщик=', 'Плательщик1=')):
                    if line.startswith('Плательщик='):
                        payer = line.replace('Плательщик=', "")
                    elif line.startswith('Плательщик1='):
                        payer = line.replace('Плательщик1=', "")
                    else:
                        payer = ""
                        print("Ошибка в Плательщике")
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
            else:
                is_income = None
                print("Ошибка в притоке/оттоке")

            payment_date = values[0]  # Дата оплаты

            # Эти значения я пока не понял как получать
            accrual_date = ""  # Дата начисления. Используется при налогах (страхование)
            article = ""  # Статья
            nds = ""  # НДС
            project = ""  # Проект

            comment = values[8]  # Комментарий

            payment_type, legal_entity, amount_in_cny, cny_exchange_rate, income, outcome, counterparty = self.get_values_depending_on_income_or_outcome(values, is_income)

            data_to_upload.append([payment_date, accrual_date, payment_type, legal_entity, article,
                                   amount_in_cny, cny_exchange_rate, income, outcome, nds,
                                   project, counterparty, comment])
        return data_to_upload

    def get_values_depending_on_income_or_outcome(self, values, is_income):
        if is_income:
            payment_type_index = 1
            legal_entity_index = 3
            rate_search_words_index = 2
            # Видимо в таблицу курс записываеться при переводе с карты ООО Модульбанк (юань),
            # а у меня, наверное, ООО Модульбанк (руб), поэтому пока поменял индексы в обратную сторону (1, 2 на 2, 1)
            counterparty_index = 4
        elif not is_income:
            payment_type_index = 2
            legal_entity_index = 4
            rate_search_words_index = 1
            counterparty_index = 3
        else:
            print("Ошибка: is_income не определен.")
            payment_type = legal_entity = amount_in_cny = cny_exchange_rate = income = outcome = counterparty = ""
            return payment_type, legal_entity, amount_in_cny, cny_exchange_rate, income, outcome, counterparty

        comment_index = 8
        sum_index = 5

        payment_type = values[payment_type_index]  # Тип оплаты (ПолучательБанк1/ПлательщикБанк1)

        if any(search_word.lower() in values[legal_entity_index].lower() for search_word in ooo_search_words):
            # search_word in string это функция, которая выполняется для всех search_word в search_words
            # any возращает true, если хоть одно значение true
            legal_entity = "ООО"  # Юр лицо (Получатель1/Плательщик1)
        elif any(search_word.lower() in values[legal_entity_index].lower() for search_word in ip_search_words):
            legal_entity = "ИП"  # Юр лицо (Получатель1/Плательщик1)
        else:
            legal_entity = ""
            print("Ошибка в Юр лице")

        # Курс CNY и Сумма в CNY
        cny_exchange_rate = ""
        amount_in_cny = ""
        # Курс указывается в комментарии. При притоке используется курс сделки, при оттоке курс ЦБ?
        if any(search_word.lower() in values[comment_index].lower() for search_word in rate_search_words):
            cny_exchange_rate, amount_in_cny = self.get_cny_exchange_rate_and_amount_in_cny(values[comment_index],
                                                                                            values[sum_index],
                                                                                            rate_search_words_index)
            # Округляем и оставляем две цифры после запятой
            cny_exchange_rate = f'{float(cny_exchange_rate):.2f}'.replace(".", ",")
            amount_in_cny = f'{float(amount_in_cny):.2f}'.replace(".", ",")

        income = ''  # Приток
        outcome = ''  # Отток
        if is_income:
            income = str(values[sum_index]).replace(".", ",")  # Приток
        elif not is_income:
            outcome = str(values[sum_index]).replace(".", ",")  # Отток

        counterparty = values[counterparty_index]  # Контрагент (Плательщик1/Получатель1)

        return payment_type, legal_entity, amount_in_cny, cny_exchange_rate, income, outcome, counterparty

    @staticmethod
    def get_cny_exchange_rate_and_amount_in_cny(comment_string, amount_in_rub, rate_search_words_index):
        # Находим конец слов "Курс сделки " или "курс ЦБ "
        rate_start_index = comment_string.find(rate_search_words[rate_search_words_index]) + len(
            rate_search_words[rate_search_words_index]) + 1

        # Находим конец числа
        for index, character in enumerate(str(comment_string[rate_start_index:])):
            if not character.isdigit():
                if not character == ("." or ","):
                    rate_end_index = rate_start_index + index
                    cny_exchange_rate = str(comment_string[rate_start_index:rate_end_index])  # Курс CNY
                    amount_in_cny = str(float(amount_in_rub.replace(",", ".")) /
                                        float(cny_exchange_rate.replace(",", ".")))  # Сумма в CNY
                    return cny_exchange_rate, amount_in_cny
        # Если конец числа не найден, т.к. строка закончилось
        cny_exchange_rate = str(comment_string[rate_start_index:])  # Курс CNY
        amount_in_cny = str(float(amount_in_rub.replace(",", ".")) /
                            float(cny_exchange_rate.replace(",", ".")))  # Сумма в CNY
        return cny_exchange_rate, amount_in_cny

    # Поиск следующей свободной строки в таблице
    @staticmethod
    def next_available_row(worksheet):
        str_list = list(filter(None, worksheet.col_values(1)))  # Берем все значения из первого столбика, потом ...?
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
            worksheet.update(f"A{self.next_available_row(worksheet)}:{chr(ord('A')-1+len(data_to_upload[0]))}"
                             f"{self.next_available_row(worksheet) + len(data_to_upload)}", data_to_upload)
            # chr(ord('A')-1+len(data_to_upload[0])) - буква алфавита по номеру начиная с заглавной A
        else:
            print("Файл не выбран или содержимое отсутствует.")
            # Вылезающее уведомление с ошибкой данных выбранного файла снизу
            if not self.data_error_snackbar:
                app = BankStatementsApp.get_running_app()
                self.data_error_snackbar = Snackbar(text="Выбранный файл отсутствует или недействителен!",
                                                 font_size=app.font_size_value,
                                                 duration=3, size_hint_x=0.8,
                                                 pos_hint={"center_x": 0.5, "center_y": 0.1})
            self.data_error_snackbar.open()


class BankStatementsApp(MDApp):
    font_size_value = "24sp"  # Размер шрифта

    def build(self):
        return MainScreen()

BankStatementsApp().run()

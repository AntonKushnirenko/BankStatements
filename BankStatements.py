from kivymd.app import MDApp
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.screen import MDScreen
from kivymd.uix.snackbar import Snackbar
import gspread
import os
from PyPDF2 import PdfReader

# Устанавливаем размер окна
from kivy.core.window import Window
window_size = (700, 525)
Window.size = window_size
# Чтобы окно нельзя было уменьшить слишком сильно
Window.minimum_width, Window.minimum_height = window_size

# Для создания exe файла
import sys
from kivy.resources import resource_add_path

google_sheet_name = "Выписки"
service_account_filename = "service_account.json"
worksheet_name = "Лист1"  # Название листа (страницы), которое выбирается снизу таблицы.
starting_directory = "/"

# Поисковые слова
ooo_search_words = ["ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО"]
ip_search_words = ["ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ", "ИП"]
rate_search_words = ["курс", "Курс сделки", "курс ЦБ"]
banking_services_search_words = ["Комиссия", "Ком-я", "Выплата начисленных процентов"]
internal_movements_search_words = ["Перевод собственных денежных средств",
                                   "Перевод средств с расчетного счета на счет 'налоговая копилка'",
                                   "Возврат средств по договору займа от учредителя",
                                   "Перевод на свою карту физ лица", "перевод на собственную карту",
                                   "перевод собственных средств", "Сбербанк Онлайн перевод"
                                   ]
withdrawal_of_money_by_the_owner_search_words = ["Сбербанк Онлайн перевод"]
communication_services_search_words = ["Билайн", "beeline", "МОРТОН ТЕЛЕКОМ", "КАНТРИКОМ"]
fuel_search_words = ["GAZPROMNEFT", "LUKOIL.AZS", "RNAZK ROSNEFT", "Газпромнефть",
                     "AZS", "АЗС", "Нефтьмагистраль", "Лукойл"]
yandex_search_words = ["ЯНДЕКС"]
ozon_search_words = ["ООО ИНТЕРНЕТ РЕШЕНИЯ"]
sbermarket_search_words = ['"МАРКЕТПЛЕЙС"', "ПАО Сбербанк"]
fraht_search_words = ["СМАРТЛОГИСТЕР"]
customs_payments_search_words = ["ФЕДЕРАЛЬНАЯ ТАМОЖЕННАЯ СЛУЖБА", "таможенного", "ТД"]
express_delivery_search_words = ["АВТОФЛОТ-СТОЛИЦА", "Деловые Линии", "Достависта",
                                 "доставки", "доставка", "грузоперевозки", "грузоперевозка", "перевозки", "перевозка"]
delivery_to_moscow_search_words = ["SHEREMETEVO-KARGO", "АВРОРА-М", "Байкал-Сервис ТК"]
accounting_search_words = ["бухгалтерских", "ДК-КОНСАЛТ"]
purchase_or_sale_search_words = ["за стяжки", "За кабельные",
                                 "Флекс", "Мегуна", "19 ДЮЙМОВ", "СВЯЗЬГАРАНТ", "Техтранссервис", "ВИССАМ",
                                 "СДС", "ВЕНЗА", "ФРЕШТЕЛ-МОСКВА", "Васильева", "СИАЙГРУПП", "МВМ", "КОНТУР-ПАК",
                                 "ЧИНЕЙКИНА", "ШАРАЕВА"]
taxes_osno_search_words = ["Казначейство России (ФНС России)", "УФК"]
warehouse_rent_search_words = ["Жилин", "Нагоркин"]
salary_fixed_search_words = ["заработная плата"]
loan_interest_repayment_search_words = ["Погашение просроч. процентов", "Оплата штрафа за проср. основной долг",
                                        "Оплата штрафа за проср. проценты", "просроч.", "проср."]
other_search_words = ["ЖИВОЙ"]
alpha_bank_search_words = ["АЛЬФА-БАНК"]
modul_bank_search_words = ["МОДУЛЬБАНК"]

# Сокращения
abbreviations = {"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ": "ООО", "МОСКОВСКИЙ ФИЛИАЛ АО КБ": "АО",
                 "СОЛДАТОВ АЛЕКСАНДР ИГОРЕВИЧ": "Солдатов А.И.", "ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ": "ИП"}

class MainScreen(MDScreen):
    def __init__(self, **kwargs):
        super(MainScreen, self).__init__(**kwargs)
        self.manager_open = False
        self.file_manager = MDFileManager(exit_manager=self.exit_manager, select_path=self.select_path)
        self.lines = []  # Строки текстового файла
        self.not_txt_snackbar = None
        self.data_error_snackbar = None
        self.is_file_selected = False
        self.format = None

    # Выбор файла
    def open_file_manager(self):
        self.file_manager.show(starting_directory)  # Здесь указываем начальную директорию

    # Этот метод будет вызываться при выборе файла
    def select_path(self, path):
        print("Путь: ", path)
        app = BankStatementsApp.get_running_app()
        # Проверяем текстовый ли файл
        if self.is_text(path):
            self.lines = [] # Очищаем строки предыдущего текстового файла
            self.format = None
            file_name = os.path.basename(path)

            if self.is_format("txt", path):
                self.format = "txt"
                # Читаем содержимое
                with open(path, "r", encoding='Windows-1251') as file:
                    self.lines = file.read().splitlines()

            # Проверяем если файл pdf
            if self.is_format("pdf", path):
                self.format = "pdf"
                # Читаем содержимое
                self.get_lines_from_pdf(path)

            # Скрываем текст kivymd кнопки и показываем текст kivy кнопки (он правильно переносит строки)
            if not self.is_file_selected:
                self.ids.select_a_file_mdbtn.icon = ''
                self.ids.select_a_file_mdbtn.text = ''
                self.ids.select_a_file_btn.opacity = 1
                self.is_file_selected = True

                # Включаем кнопку выгрузки файла в Google Таблицу
                self.ids.upload_a_file_btn.disabled = False

            # Меняем текст на кнопке выбора файла на имя выбранного файла
            self.ids.select_a_file_btn.text = f"[color=#696969]{file_name}[/color]"

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
    def is_text(path):
        if path.endswith(('.txt', '.pdf')):
            return True
        else:
            print("Неподдерживаемый формат файла")
            return False

    # Проверяем формат файла
    @staticmethod
    def is_format(format, path):
        format_ending = f'.{format}'
        if path.endswith(format_ending):
            return True

    def exit_manager(self, instance):
        self.manager_open = False
        self.file_manager.close()

    # Получаем содержимое pdf файла в виде строчек
    def get_lines_from_pdf(self, path):
        reader = PdfReader(path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        print(text)
        self.lines = text.split("\n")
        print(self.lines)

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

                # Если ДатаСписано или ДатаПоступило позже Дата, берется ДатаСписано или ДатаПоступило
                if line.startswith('ДатаСписано='):
                    date_write_off_or_received = line.replace('ДатаСписано=', "")
                elif line.startswith('ДатаПоступило='):
                    date_write_off_or_received = line.replace('ДатаПоступило=', "")
                    if date_write_off_or_received != "":
                        if dates:
                            if dates[-1] != date_write_off_or_received:
                                dates[-1] = date_write_off_or_received

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

    # Получаем значения из pdf файла, необходимые для вычисления данных, которые потом будем вносить в таблицу
    def get_required_values_from_pdf(self):
        dates = []  # Дата
        operation_names = []  # Название операции
        sums_of_money = []  # Сумма

        is_sberbank = False
        is_document_section_started = False
        for index, line in enumerate(self.lines):
            if line.startswith('Сформировано в СберБанк Онлайн'):
                is_sberbank = True
            if is_sberbank:
                if line.startswith('Расшифровка операций'):
                    is_document_section_started = True

                if is_document_section_started:
                    # Проверяем подходит ли строка под формат даты
                    if all(character.isnumeric() for character in [line[:2], line[3:5], line[6:10], line[11:13], line[14:16]]) and all(character == "." for character in [line[2], line[5]]) and line[13]==":":
                        date = line[:10]
                        print(line)
                        dates.append(date)

                        # Относительно строки даты рассматриваем следующие строки
                        operation_names.append(self.lines[index + 1][17:])

                        amount_of_money_line = self.lines[index + 2].replace("\xa0", "")
                        if amount_of_money_line.rfind("-") >= 0:
                            amount_beginning_index = amount_of_money_line.rfind("-")
                        else:
                            for character in reversed(amount_of_money_line):
                                if not character.isnumeric():
                                    if not character == ",":
                                        amount_beginning_index = amount_of_money_line.rindex(character)+1
                                        break

                        amount_of_money = amount_of_money_line[amount_beginning_index:]
                        sums_of_money.append(amount_of_money)

        required_values = list(map(list, zip(dates, operation_names, sums_of_money)))
        print(required_values)
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
            nds = ""  # НДС 
            project = ""  # Проект

            comment = values[8]  # Комментарий

            payment_type, legal_entity, article, amount_in_cny, cny_exchange_rate, income, outcome, counterparty = self.get_values_depending_on_income_or_outcome(
                values, is_income)

            data_to_upload.append([payment_date, accrual_date, payment_type, legal_entity, article,
                                   amount_in_cny, cny_exchange_rate, income, outcome, nds,
                                   project, counterparty, comment])
        return data_to_upload

    def get_data_to_upload_from_pdf(self, required_values):
        data_to_upload = []
        for values in required_values:
            payment_date = values[0]  # Дата оплаты

            # Эти значения я пока не понял как получать
            accrual_date = ""  # Дата начисления. Используется при налогах (страхование)
            nds = ""  # НДС
            project = ""  # Проект

            article = self.get_article(values[1], values[1], not values[2].startswith("-"))

            comment = ""  # Комментарий
            if values[1].startswith('Сбербанк Онлайн перевод'):
                comment = values[1]

            payment_type = "Личная карта сбербанк"  # Нужна проверка, что действительно сбербанк, а не hardcoded значение
            legal_entity = "ИП"

            amount_in_cny = ""
            cny_exchange_rate = ""
            income = ""
            outcome = ""
            if values[2].startswith("-"):
                outcome = values[2][1:]
            else:
                income = values[2]
            counterparty = ""
            if not values[1].startswith('Сбербанк Онлайн перевод'):
                counterparty = values[1]



            data_to_upload.append([payment_date, accrual_date, payment_type, legal_entity, article,
                                   amount_in_cny, cny_exchange_rate, income, outcome, nds,
                                   project, counterparty, comment])
        return data_to_upload

    def get_values_depending_on_income_or_outcome(self, values, is_income):
        if is_income:
            bank_name_type_index = 1
            legal_entity_index = 3
            rate_search_words_index = 2
            # Видимо в таблицу курс записываеться при переводе с карты ООО Модульбанк (юань),
            # а у меня, наверное, ООО Модульбанк (руб), поэтому пока поменял индексы в обратную сторону (1, 2 на 2, 1)
            counterparty_index = 4
        elif not is_income:
            bank_name_type_index = 2
            legal_entity_index = 4
            rate_search_words_index = 1
            counterparty_index = 3
        else:
            print("Ошибка: is_income не определен.")
            payment_type = legal_entity = amount_in_cny = cny_exchange_rate = income = outcome = counterparty = ""
            return payment_type, legal_entity, amount_in_cny, cny_exchange_rate, income, outcome, counterparty

        comment_index = 8
        sum_index = 5

        '''
        # recipient_index = 4

        # Тип оплаты (ПолучательБанк1/ПлательщикБанк1)
        # Нужно еще доделать подписи валюты в скобочках: (руб.), (юань), (usd)
        if any(search_word.lower() in values[legal_entity_index].lower() for search_word in ooo_search_words):
            # recipient_index = 4 - индекс Плательщик1, а мне всегда нужен Плательщик1
            # Как оказалось - не всегда. kl_to_1c (8).txt: 05.06.2023 13605,00 Плательщик1 - ООО Яндекс,
            # а должно быть ИП "Модуль банк", следовательно брать надо из
            # Получатель1=Индивидуальный предприниматель Солдатов Александр Игоревич
            payment_type_legal_entity = "ООО"
            # Здесь нужно поменять на ИП/ООО Альфа-банк/Модуль-банк
        elif any(search_word.lower() in values[legal_entity_index].lower() for search_word in ip_search_words):
            payment_type_legal_entity = "ИП"
        # Видимо еще нужна подпись "Личная карта" и "Налоговая копилка"
        else:
            payment_type_legal_entity = ""
            print("Ошибка в Юр лице для типа оплаты")

        if any(search_word.lower() in values[bank_name_type_index].lower() for search_word in alpha_bank_search_words):
            payment_type = f'{payment_type_legal_entity} Альфа'
        elif any(
                search_word.lower() in values[bank_name_type_index].lower() for search_word in modul_bank_search_words):
            payment_type = f'{payment_type_legal_entity} Модульбанк'
        else:
            payment_type = f'{payment_type_legal_entity} {values[bank_name_type_index]}'
        '''

        # Юр лицо (Получатель1/Плательщик1)
        if any(search_word.lower() in values[legal_entity_index].lower() for search_word in ooo_search_words):
            # search_word in string это функция, которая выполняется для всех search_word в search_words
            # any возращает true, если хоть одно значение true
            legal_entity = "ООО"
        elif any(search_word.lower() in values[legal_entity_index].lower() for search_word in ip_search_words):
            legal_entity = "ИП"
        else:
            legal_entity = ""
            print("Ошибка в Юр лице")

        # Сокращение строк с названиями банков
        if any(search_word.lower() in values[bank_name_type_index].lower() for search_word in alpha_bank_search_words):
            payment_type = f'{legal_entity} Альфа'  # Еще бывает "Личная карта Альфа"
        elif any(
                search_word.lower() in values[bank_name_type_index].lower() for search_word in modul_bank_search_words):
            payment_type = f'{legal_entity} Модульбанк'
        else:
            payment_type = f'{legal_entity} {values[bank_name_type_index]}'

        # Проверка на тип отплаты "Налоговая копилка"
        if self.is_tax_piggy_bank(values[comment_index], is_income):
            payment_type = "Налоговая копилка"

        article = self.get_article(values[comment_index], values[counterparty_index], is_income)  # Статья

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

        # Сокращаем Юр лица
        abbreviated_counterparty = None
        for full, abbreviated in abbreviations.items():
            if not abbreviated_counterparty:
                abbreviated_counterparty = self.counterparty_case_insensitive_replace(counterparty, full, abbreviated)

        # Присваеваем сокращенное значение, если сокращение нашлось
        if abbreviated_counterparty:
            counterparty = abbreviated_counterparty

        # Если изначально у контрагента было "(ИП)" - ставим "ИП" в начало
        if "(ИП)" in counterparty:
            counterparty = counterparty.replace("(ИП)", "")
            counterparty = "ИП " + counterparty

        return payment_type, legal_entity, article, amount_in_cny, cny_exchange_rate, income, outcome, counterparty

    # Заменяем юр лица в строке независимо от заглавных или маленьких букв
    @staticmethod
    def counterparty_case_insensitive_replace(counterparty_string, string_to_replace, replacement_string):
        if counterparty_string:
            if string_to_replace.upper() in counterparty_string.upper():
                start_index = counterparty_string.upper().find(string_to_replace.upper())
                end_index = start_index + len(string_to_replace)
                return counterparty_string[:start_index] + replacement_string + counterparty_string[end_index:]

    # Проверка на тип отплаты "Налоговая копилка"
    @staticmethod
    def is_tax_piggy_bank(comment_string, is_income):
        if "налоговая копилка" in comment_string.lower() and is_income:
        # Если есть фраза 'налоговая копилка' в комментарии и идет приток денег
            return True

    # Получаем курс юаней и сумму в юанях
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

    # Получение статьи через поисковые слова в комментариях или контрагенте
    @staticmethod
    def get_article(comment_string, counterparty_string, is_income):
        if comment_string or counterparty_string:
            if any(search_word.lower() in comment_string.lower() for search_word in banking_services_search_words):
                return "Банковское обслуживание и комиссии"
            elif any(search_word.lower() in comment_string.lower() for search_word in rate_search_words):
                return "Обмен валют"
            elif any(search_word.lower() in comment_string.lower() for search_word in withdrawal_of_money_by_the_owner_search_words):
                if not is_income:
                    return "Вывод ДС собственником"  # Не обязательно
                elif is_income:
                    return "Внутренние перемещения"  # Не обязательно
            elif any(search_word.lower() in comment_string.lower() for search_word in internal_movements_search_words):
                return "Внутренние перемещения"
            elif any(search_word.lower() in (" ".join([comment_string.lower(), counterparty_string.lower()]))
                     for search_word in communication_services_search_words):
                return "Услуги связи и Интернет"
            elif any(search_word.lower() in (" ".join([comment_string.lower(), counterparty_string.lower()]))
                     for search_word in fuel_search_words):
                return "Топливо"
            elif any(search_word.lower() in (" ".join([comment_string.lower(), counterparty_string.lower()]))
                     for search_word in customs_payments_search_words):
                return "Таможенные платежи"
            elif any(search_word.lower() in counterparty_string.lower()
                     for search_word in delivery_to_moscow_search_words):
                return "Доставка до МСК"
            elif any(search_word.lower() in (" ".join([comment_string.lower(), counterparty_string.lower()]))
                     for search_word in express_delivery_search_words):
                return "Курьерская доставка"
            elif any(search_word.lower() in (" ".join([comment_string.lower(), counterparty_string.lower()]))
                     for search_word in accounting_search_words):
                return "Бухгалтерия"
            elif any(search_word.lower() in (" ".join([comment_string.lower(), counterparty_string.lower()]))
                     for search_word in purchase_or_sale_search_words):
                if is_income:
                    return "Оптовые продажи"
                elif not is_income:
                    return "Закупка товара"
            elif any(search_word.lower() in counterparty_string.lower() for search_word in taxes_osno_search_words):
                return "Налоги ОСНО"
            elif any(search_word.lower() in counterparty_string.lower() for search_word in warehouse_rent_search_words):
                return "Аренда склада"
            elif any(search_word.lower() in comment_string.lower() for search_word in salary_fixed_search_words):
                return "Зарплата - фикс"
            elif any(search_word.lower() in comment_string.lower() for search_word in loan_interest_repayment_search_words):
                return "Погашение процентов по кредиту"
            elif any(search_word.lower() in counterparty_string.lower() for search_word in yandex_search_words):
                return "Я.Маркет"
            elif any(search_word.lower() in counterparty_string.lower() for search_word in ozon_search_words):
                return "Ozon"
            elif any(search_word.lower() in counterparty_string.lower() for search_word in sbermarket_search_words):
                return "Сбермаркет"
            elif any(search_word.lower() in counterparty_string.lower() for search_word in fraht_search_words):
                return "Фрахт"  # ООО СМАРТЛОГИСТЕР не только фрахт, но и может быть Таможенные платежи
            elif any(search_word.lower() in counterparty_string.lower() for search_word in other_search_words):
                return "Прочее"
            else:
                return ""
        else:
            return ""

    # Поиск следующей свободной строки в таблице
    @staticmethod
    def next_available_row(worksheet):
        str_list = list(filter(None, worksheet.col_values(1)))  # Берем все значения из первого столбика, потом ...?
        return len(str_list) + 1

    # Работа с гугл таблицами
    def upload_to_googledrive(self):
        service_account = gspread.service_account(filename=BankStatementsApp.resource_path(service_account_filename))
        spreadsheet = service_account.open(google_sheet_name)
        worksheet = spreadsheet.worksheet(worksheet_name)

        if self.format == "txt":
            required_values = self.get_required_values()
            income_checking_account = self.get_income_checking_account()  # РасчСчет
            data_to_upload = self.get_data_to_upload(required_values, income_checking_account)
        elif self.format == "pdf":
            required_values = self.get_required_values_from_pdf()
            data_to_upload = self.get_data_to_upload_from_pdf(required_values)
        if data_to_upload:  # Проверяем не пустой ли файл
            worksheet.update(f"A{self.next_available_row(worksheet)}:{chr(ord('A') - 1 + len(data_to_upload[0]))}"
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
    font_size_value = "25sp"  # Размер шрифта

    def build(self):
        # Устанавливаем название и иконку окна приложения
        self.title = 'Выгрузка банковских выписок в Google Таблицу "АРТЕЛЬ/ФИНАНСЫ"'
        self.icon = self.resource_path('artel_icon.png')
        return MainScreen()

    # Для создания exe файла
    # Функция возвращает путь к файлам, чтобы в exe отображались картинки и находился файл service_account.json
    @staticmethod
    def resource_path(relative_path):
        # returns an absolute path
        try:
            base_path = sys._MEIPASS
        except Exception:
            # derive the absolute path of the relative_path
            base_path = os.path.abspath('.')
        return os.path.join(base_path, relative_path)

BankStatementsApp().run()

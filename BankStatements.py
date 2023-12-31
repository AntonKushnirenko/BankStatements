from kivymd.app import MDApp

from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.screen import MDScreen
from kivymd.uix.snackbar import Snackbar
from kivymd.uix.dialog import MDDialog
from kivymd.uix.button import MDFlatButton
from kivymd.uix.recycleview import MDRecycleView

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

# Для сохранения настроек
from kivy.storage.jsonstore import JsonStore
store = JsonStore("settings.json")

# Информация о таблице
google_sheet_name = "АРТЭЛЬ/ База данных"  # "АРТЭЛЬ/ ФИНАНСЫ"
google_sheet_name_default = google_sheet_name  # Значение на случай сброса настроек
service_account_filename = "service_account.json"
worksheet_name = "БД ДДС"  # Название листа (страницы), которое выбирается снизу таблицы. "БД ДДС"
worksheet_name_default = worksheet_name
# Начальная директория файлового менеджера
starting_directory = "/"
starting_directory_default = starting_directory

is_cny_statement = False  # Является ли выписка со счета в юанях
is_cny_statement_manually = False  # Значение "является ли выписка со счета в юанях", установленное вручную в настройках
show_article_options = False  # Настройка "показывать варианты статей через "/" или нет"

# Получаем название переменной в строке
def get_variable_name(variable):
    for name in globals():
        if id(globals()[name]) == id(variable):
            return name

# Получаем сохранненные в настройках значения
settings_items = [google_sheet_name, worksheet_name, starting_directory]
for item in settings_items:
    item_name = get_variable_name(item)
    if store.exists(item_name):
        globals()[item_name] = store.get(item_name)[item_name]

# Отдельно для True/False значений
if store.exists('is_cny_statement_manually'):
    is_cny_statement_manually = store.get('is_cny_statement_manually')['is_cny_statement_manually']
if store.exists('show_article_options'):
    show_article_options = store.get('show_article_options')['show_article_options']

# Поисковые слова
ooo_search_words = ("ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО")
ip_search_words = ("ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ", "ИП", "СОЛДАТОВ АЛЕКСАНДР ИГОРЕВИЧ")
rate_search_words = ("курс", "Курс сделки", "курс ЦБ")
banking_services_search_words = ("Комиссия", "Ком-я", "Выплата начисленных процентов")
internal_movements_search_words = ("Перевод собственных денежных средств",
                                   "Перевод средств с расчетного счета на счет 'налоговая копилка'",
                                   "Возврат средств по договору займа от учредителя",
                                   "Перевод на свою карту", "перевод на собственную карту",
                                   "перевод собственных средств", "Перевод остатка денежных средств",
                                   "Сбербанк Онлайн перевод", "Прочие выплаты", "ATM", "Тинькофф Банк",
                                   "возврат долга по договору займа", "Перевод на свой счет в дргуой банк",
                                   "перевод своих средств", "взнос по договору займа",
                                   "Возврат займа единственному учредителю",)
withdrawal_of_money_by_the_owner_search_words = ("Сбербанк Онлайн перевод",
                                                 "PYATEROCHKA", "QSR", "APTEKA", "RPS SERVIS", "SOKOLOV", "CHaihona",
                                                 "LAWAVE", "EUROAUTO", "TC SKHODNENSKAYA", "MAPP_SBERBANK_ONL@IN_PAY",
                                                 "IP SAMOTAEVA A.Z.", "AVTOTREYD", "VKUSVILL", "IP Kuznetsov",
                                                 "COFFEE", "AVTODOR", "IP LESONEN M.A.", "IP SAMSONOV V.N.",
                                                 "FEJS TU FEJS", "UMNYASHKA-SCHOOL.RU", "KORCHMA TARAS BULBA",
                                                 "P2P_byPhone_tinkoff-bank", "PEREKRESTOK", "IP BANYAS P.V.",
                                                 "PARKOVKA", "MOSKALENKO", "PE MOSKALENKO ANDREY", "REPETICIONNAYABAZAT",
                                                 "ZOOMAGAZIN", "MYASNAYA LAVKA", "OKEY", "РЕСО-Гарантия", "SITI LAJF",
                                                 "PRAZHSKAYA", "MAGBURGER", "TEREMOK", "IP Badalyan LR", "owl Bean",
                                                 "FISHBAZAAR", "RYNOK", "SUPERMARKET", "PRODUKTY", "BI-BI",
                                                 "IP LYASHCHENKO L.V.", "WWW.1PARTS.RU", "IP SHASHKIN O V",
                                                 # IP SHASHKIN O V - внутренние перемещения, но я не уверен, что это так должно быть
                                                 "VSEINSTRUMENTI.RU")  # VSEINSTRUMENTI.RU - прочее при притоке (2223,00 07.06)
contribution_of_money_by_the_owner_search_words = ("Перевод собственных средств",  # Внесение ДС собственником
                                                   "Перевод средств с расчетного счета на счет 'налоговая копилка'",
                                                   "Перевод остатка денежных средств", )
returns_income_search_words = ("ВОЗВРАТ ОШИБОЧНО ПЕРЕЧИСЛЕННЫХ СРЕДСТВ", "возврат")  # Возвраты - приток
caching_search_words = ("СТАЛЬНОЕ СЕРДЦЕ", 'ООО "ТЕХКОМ"')
marketing_search_words = ("продвижению", )
providing_bidding_search_words = ("обеспечение заявки", )  # Обеспечение торгов
education_search_words = ("за обучение", "за образовательные услуги", "ЦЗП")
sdek_search_words = ("СДЭК", )
insurance_contributions_with_salary = ("страхования", )  # Стховые взносы с ЗП
# Две идентичные строчки на 52 рубля (за июнь и август) имеют разные статьи: одна Налоги ОСНО, другая Страховые взнносы
packaging_search_words = ("упаков", )  # Упаковка
wildberries_search_words = ("ВАЙЛДБЕРРИЗ", )
communication_services_search_words = ("Билайн", "beeline", "МОРТОН ТЕЛЕКОМ", "КАНТРИКОМ")
fuel_search_words = ("GAZPROMNEFT", "LUKOIL.AZS", "RNAZK ROSNEFT", "Газпромнефть",
                     "AZS", "АЗС", "Нефтьмагистраль", "Лукойл")
yandex_search_words = ("ЯНДЕКС", )
ozon_search_words = ("ООО ИНТЕРНЕТ РЕШЕНИЯ", 'ООО "ИНТЕРНЕТ РЕШЕНИЯ"')
sbermarket_search_words = ('"МАРКЕТПЛЕЙС"', "ПАО Сбербанк")
fraht_search_words = ("СМАРТЛОГИСТЕР", "транспортные услуги")
customs_payments_search_words = ("ФЕДЕРАЛЬНАЯ ТАМОЖЕННАЯ СЛУЖБА", "таможенного", "ТД", "ДТ")
express_delivery_search_words = ("АВТОФЛОТ-СТОЛИЦА", "Деловые Линии", "Достависта",
                                 "доставки", "доставка", "грузоперевозки", "грузоперевозка", "перевозки", "перевозка")
delivery_to_moscow_search_words = ("SHEREMETEVO-KARGO", "АВРОРА-М", "Байкал-Сервис ТК")
accounting_search_words = ("бухгалтерских", "ДК-КОНСАЛТ", "Диадок")
purchase_or_sale_search_words = ("за стяжки", "за стяжку", "За кабельные", "за измельчитель", "за электрооборудование",
                                 "за электротовары", "за хомуты", "ЗА ТМЦ", "за строй материалы", "за этикетки",
                                 "за фасадные комплектующие", "ЗА ОБОРУДОВАНИЕ", "за розетки", "за кабель",
                                 "За крепеж", "блоки розеток", "За разъемы", "За Щит ЩРН", "за выключатели",
                                 "за нейлоновые стяжки", "за материалы", "за квадрокоптер",
                                 "Флекс", "Мегуна", "19 ДЮЙМОВ", "СВЯЗЬГАРАНТ", "Техтранссервис", "ВИССАМ",
                                 "СДС", "ВЕНЗА", "ФРЕШТЕЛ-МОСКВА", "Васильева", "СИАЙГРУПП", "МВМ", "КОНТУР-ПАК",
                                 "ЧИНЕЙКИНА", "ШАРАЕВА", "АВАЛОН", "МОСКАБЕЛЬ", "ЗАВОД ТРУД",
                                 "ТИМОФЕЕВ", "БОРОДИН", "Приль", )
taxes_osno_search_words = ("Казначейство России (ФНС России)", "УФК")
taxes_usn_search_words = ("Налог при упрощенной системе налогообложения", "усн")
warehouse_rent_search_words = ("Жилин", "Нагоркин")
salary_fixed_search_words = ("заработная плата", "У. НАТАЛЬЯ ВАЛЕРЬЕВНА", "С. ДМИТРИЙ ВАДИМОВИЧ",
                             "Для зачисления на счет Солдатова Александра Игоревича аванс", )
loan_interest_repayment_search_words = ("Погашение просроч. процентов", "Оплата штрафа за проср. основной долг",
                                        "Оплата штрафа за проср. проценты", "просроч.", "проср.")
other_search_words = ("ЖИВОЙ", "АДВОКАТ", "РСЦ", )
# Прочее почему-то часто "возврат", но я не стал добавлять, ведь есть "Возврат - приток"
alpha_bank_search_words = ("АЛЬФА-БАНК", )
modul_bank_search_words = ("МОДУЛЬБАНК", )
sberbank_comments_search_words = ("P2P_byPhone_tinkoff-bank", "ATM", "Тинькофф Банк", "oplata_beeline",
                                  "Прочие выплаты")

# Сокращения
abbreviations = {"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ": "ООО", "МОСКОВСКИЙ ФИЛИАЛ АО КБ": "АО",
                 "СОЛДАТОВ АЛЕКСАНДР ИГОРЕВИЧ": "Солдатов А.И.", "ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ": "ИП",
                 "НЕКОММЕРЧЕСКАЯ ОРГАНИЗАЦИЯ": "НКО", "Закрытое акционерное общество": "ЗАО",
                 "У. НАТАЛЬЯ ВАЛЕРЬЕВНА": "Наталья Валерьевна У"}
# Статьи
articles_by_search_words_for_comments = {banking_services_search_words: "Банковское обслуживание и комиссии",
                                         rate_search_words: "Обмен валют",
                                         salary_fixed_search_words: "Зарплата - фикс",
                                         internal_movements_search_words: "Внутренние перемещения",
                                         loan_interest_repayment_search_words: "Погашение процентов по кредиту",
                                         taxes_usn_search_words: "Налоги УСН",
                                         marketing_search_words: "Маркетинг",
                                         providing_bidding_search_words: "Обеспечение торгов",
                                         packaging_search_words: "Упаковка"}
articles_by_search_words_for_counterpartys = {delivery_to_moscow_search_words: "Доставка до МСК",
                                              taxes_osno_search_words: "Налоги ОСНО",
                                              warehouse_rent_search_words: "Аренда склада",
                                              yandex_search_words: "Я.Маркет",
                                              ozon_search_words: "Ozon",
                                              sbermarket_search_words: "Сбермаркет",
                                              other_search_words: "Прочее",
                                              wildberries_search_words: "Wildberries",
                                              sdek_search_words: "СДЭК",
                                              caching_search_words: "Кэширование",
                                              insurance_contributions_with_salary: "Страховые взносы с ЗП"}
articles_by_search_words_for_comments_or_counterpartys = {communication_services_search_words: "Услуги связи и Интернет",
                                                          fuel_search_words: "Топливо",
                                                          customs_payments_search_words: "Таможенные платежи",
                                                          express_delivery_search_words: "Курьерская доставка",
                                                          accounting_search_words: "Бухгалтерия",
                                                          fraht_search_words: "Фрахт",
                                                          education_search_words: "Обучение"}
articles_by_search_words_for_comments_if_income = {withdrawal_of_money_by_the_owner_search_words: "Внутренние перемещения",
                                                   returns_income_search_words: "Возвраты - приток",
                                                   contribution_of_money_by_the_owner_search_words: "Внесение ДС собственником"}
articles_by_search_words_for_comments_if_outcome = {withdrawal_of_money_by_the_owner_search_words: "Вывод ДС собственником"}
articles_by_search_words_for_comments_or_counterpartys_if_income = {purchase_or_sale_search_words: "Оптовые продажи"}
articles_by_search_words_for_comments_or_counterpartys_if_outcome = {purchase_or_sale_search_words: "Закупка товара"}

priority_order_of_articles = ("Зарплата - фикс", "Оптовые продажи", "Налоги УСН", "Налоги ОСНО",
                              "Вывод ДС собственником", "Таможенные платежи", "СДЭК", "Кэширование",
                              "Возвраты - приток", "Маркетинг", )


class MainScreen(MDScreen):
    def __init__(self, **kwargs):
        super(MainScreen, self).__init__(**kwargs)
        self.manager_open = False
        self.file_manager = MDFileManager(exit_manager=self.exit_manager, select_path=self.select_path)
        self.lines = []  # Строки текстового файла
        self.not_txt_snackbar = None
        self.data_error_snackbar = None
        self.upload_error_snackbar = None
        self.not_enough_rows_error_snackbar = None
        self.google_sheet_name_error_snackbar = None
        self.worksheet_name_error_snackbar = None
        self.is_file_selected = False
        self.format = None
        self.settings_dialog = None  # Меню с настройками

    # Выбор файла
    def open_file_manager(self):
        self.file_manager.show(starting_directory)  # Здесь указываем начальную директорию

    # Этот метод будет вызываться при выборе файла
    def select_path(self, path):
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
        self.lines = text.split("\n")

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

        global is_cny_statement
        global is_cny_statement_manually
        # Если в настройках вручную выбрано, что выписка в юанях - устанавливаем, что выписка в юанях.
        # Если нет - обнуляем значение перед проверкой
        is_cny_statement = is_cny_statement_manually

        is_document_section_started = False
        for line in self.lines:
            if line.startswith('СекцияДокумент='):
                is_document_section_started = True

            if is_document_section_started:
                # Предварительная проверка в юанях ли выписка, путем поиска "CNY" в НазначениеПлатежа
                if not is_cny_statement:
                    is_cny_statement = self.check_if_statement_in_cny(line)

                if line.startswith('Дата='):
                    date = line.replace('Дата=', "")
                    dates.append(date)

                # Если ДатаСписано или ДатаПоступило позже Дата, берется ДатаСписано или ДатаПоступило
                date_write_off_or_received = ""
                if line.startswith('ДатаПоступило='):
                    date_write_off_or_received = line.replace('ДатаПоступило=', "")
                elif line.startswith('ДатаСписано='):
                    date_write_off_or_received = line.replace('ДатаСписано=', "")
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
                        dates.append(date)

                        # Проверяем, что строка с названием операции не разделилась на две
                        while not any(character.isnumeric() for character in self.lines[index + 2]):
                            self.lines[index + 1] += f" {self.lines[index + 2]}"
                            self.lines.pop(index +2)

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
        return required_values, is_sberbank

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

            # Эти значения пока никак не получаются
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

    # Формируем данные которые будем уже непосредственно вносить в таблицу из pdf
    def get_data_to_upload_from_pdf(self, required_values, is_sberbank):
        data_to_upload = []
        for values in required_values:
            payment_date = values[0]  # Дата оплаты

            # Эти значения пока никак не получаются
            accrual_date = ""  # Дата начисления. Используется при налогах (страхование)
            nds = ""  # НДС
            project = ""  # Проект

            article = self.get_article(values[1], values[1], not values[2].startswith("-"))

            # Убираем приписки "*Название города* RUS" в конце названия операции
            if values[1].endswith("RUS"):
                city_name_first_word = values[1].split()[-2]
                # Если между предыдущими двумя словами один пробел - название города в два слова
                if values[1][values[1].find(values[1].split()[-2]) - 1].isspace() and not values[1][
                    values[1].find(values[1].split()[-2]) - 2].isspace():
                    city_name_second_word = values[1].split()[-3]
                    values[1] = values[1].replace(city_name_second_word, " ")
                values[1] = values[1].replace("RUS", " ")
                values[1] = values[1].replace(city_name_first_word, " ")

            counterparty = ""
            if not values[1].startswith('Сбербанк Онлайн перевод'):
                # Проверяем отдельные случаи, когда значение должно быть в комментарии, а не в контрагенте
                if not any(search_word.lower() in values[1].lower() for search_word in sberbank_comments_search_words):
                    counterparty = values[1]

            comment = ""  # Комментарий
            if any(search_word.lower() in values[1].lower() for search_word in sberbank_comments_search_words):
                comment = values[1]

            if values[1].startswith('Сбербанк Онлайн перевод'):
                # Проверяем есть ли имя человека в строке 'Сбербанк Онлайн перевод'
                for index, character in enumerate(reversed(values[1])):
                    if character.isnumeric():
                        last_number_index = len(values[1]) - 1 - index
                        if last_number_index + 1 != len(values[1]):
                            name_beginning_index = last_number_index + 2
                            name = values[1][name_beginning_index:]
                            if name != "":
                                # Если есть - записываем в контрагента и убираем из комментария
                                counterparty = name
                                comment = values[1].replace(name, "")
                            break
                        else:
                            # Если нет - оставляем комментарий 'Сбербанк Онлайн перевод...'
                            comment = values[1]
                            break

            # Сокращаем Юр лица в контрагентах
            if any(full_phrase.upper() in counterparty.upper() for full_phrase in abbreviations.keys()):
                counterparty = self.abbreviaton(counterparty)

            # Проверяем, что тип оплаты действительно Сбербанк
            if is_sberbank:
                payment_type = "Личная карта сбербанк"
            legal_entity = "ИП"

            amount_in_cny = ""
            cny_exchange_rate = ""
            income = ""
            outcome = ""
            if values[2].startswith("-"):
                outcome = values[2][1:]
            else:
                income = values[2]

            data_to_upload.append([payment_date, accrual_date, payment_type, legal_entity, article,
                                   amount_in_cny, cny_exchange_rate, income, outcome, nds,
                                   project, counterparty, comment])
        return data_to_upload

    def get_values_depending_on_income_or_outcome(self, values, is_income):
        global is_cny_statement

        if is_income:
            bank_name_type_index = 1
            legal_entity_index = 3
            rate_search_words_index = 1  # Поменял, чтобы везде был курс сделки (курс ЦБ не используется)
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

        # Добавление (юань) к концу типа оплаты, если выписка со счета в юанях
        if is_cny_statement:
            payment_type += " (юань)"

        # Проверка на тип отплаты "Налоговая копилка"
        if self.is_tax_piggy_bank(values[comment_index], is_income):
            payment_type = "Налоговая копилка"

        article = self.get_article(values[comment_index], values[counterparty_index], is_income)  # Статья

        # Курс CNY и Сумма в CNY
        cny_exchange_rate = ""
        amount_in_cny = ""
        amount_in_rub = ""
        # Курс указывается в комментарии.
        if any(search_word.lower() in values[comment_index].lower() for search_word in rate_search_words):
            cny_exchange_rate = self.get_cny_exchange_rate(values[comment_index], rate_search_words_index)
            # Если выписка уже в юанях
            if is_cny_statement:
                amount_in_rub = self.get_amount_in_rub(values[comment_index], values[sum_index], cny_exchange_rate)
                amount_in_rub = f'{float(amount_in_rub):.2f}'.replace(".", ",")
            # Если выписка не в юанях (в рублях)
            else:
                amount_in_cny = self.get_amount_in_cny(values[comment_index], values[sum_index], cny_exchange_rate)
                amount_in_cny = f'{float(amount_in_cny):.2f}'.replace(".", ",")
            # Округляем и оставляем две цифры после запятой
            cny_exchange_rate = f'{float(cny_exchange_rate):.2f}'.replace(".", ",")

        income = ''  # Приток
        outcome = ''  # Отток
        if is_income:
            # Если выписка в юанях - в приток записываем переведенное в рубли значение
            if is_cny_statement:
                income = amount_in_rub
                amount_in_cny = str(values[sum_index]).replace(".", ",")
            else:
                income = str(values[sum_index]).replace(".", ",")  # Приток
        elif not is_income:
            if is_cny_statement:
                outcome = amount_in_rub
                amount_in_cny = str(values[sum_index]).replace(".", ",")
            else:
                outcome = str(values[sum_index]).replace(".", ",")  # Отток

        counterparty = values[counterparty_index]  # Контрагент (Плательщик1/Получатель1)

        # Сокращаем Юр лица в контрагентах
        if any(full_phrase.upper() in counterparty.upper() for full_phrase in abbreviations.keys()):
            counterparty = self.abbreviaton(counterparty)

        # Убираем расченый счет из контрагента
        if "Р/С" in counterparty:
            # Если перед Р/С стоит пробел - удаляем его тоже
            if counterparty[counterparty.index("Р/С")-1] == " ":
                counterparty = counterparty[:(counterparty.index("Р/С")-1)]
            else:
                counterparty = counterparty[:counterparty.index("Р/С")]

        # Если изначально у контрагента было "(ИП)" - ставим "ИП" в начало
        if "(ИП)" in counterparty:
            counterparty = counterparty.replace("(ИП)", "")
            counterparty = "ИП " + counterparty

        return payment_type, legal_entity, article, amount_in_cny, cny_exchange_rate, income, outcome, counterparty

    # Сокращаем Юр лица в контрагентах
    def abbreviaton(self, counterparty):
        for full, abbreviated in abbreviations.items():
            abbreviated_counterparty = self.counterparty_case_insensitive_replace(counterparty, full, abbreviated)
            if abbreviated_counterparty:
                return abbreviated_counterparty

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

    # Получаем курс юаней
    @staticmethod
    def get_cny_exchange_rate(comment_string, rate_search_words_index):
        # Находим конец слов "Курс сделки " или "курс ЦБ "
        rate_start_index = comment_string.find(rate_search_words[rate_search_words_index]) + len(
            rate_search_words[rate_search_words_index]) + 1

        # Находим конец числа
        for index, character in enumerate(str(comment_string[rate_start_index:])):
            if not character.isdigit():
                if not character == ("." or ","):
                    rate_end_index = rate_start_index + index
                    cny_exchange_rate = str(comment_string[rate_start_index:rate_end_index])  # Курс CNY
                    return cny_exchange_rate
        # Если конец числа не найден, т.к. строка закончилось
        cny_exchange_rate = str(comment_string[rate_start_index:])  # Курс CNY
        return cny_exchange_rate

    # Получаем сумму в юанях из рублей
    @staticmethod
    def get_amount_in_cny(comment_string, amount_in_rub, cny_exchange_rate):
        amount_in_cny = str(float(amount_in_rub.replace(",", ".")) /
                            float(cny_exchange_rate.replace(",", ".")))  # Сумма в CNY
        return amount_in_cny

    # Получаем сумму в рублях из юаней
    @staticmethod
    def get_amount_in_rub(comment_string, amount_in_cny, cny_exchange_rate):
        amount_in_rub = str(float(amount_in_cny.replace(",", ".")) *
                            float(cny_exchange_rate.replace(",", ".")))  # Сумма в RUB
        return amount_in_rub

    # Проверяем в юанях ли выписка, путем поиска "CNY" в НазначениеПлатежа
    # Но возможно, что комментария с "CNY" может не быть в выписке,
    # которая является выпиской в юанях, поэтому добавлены настройки, где можно указать, что выписка в юанях вручную
    def check_if_statement_in_cny(self, line):
        if line.startswith('НазначениеПлатежа='):
            if "CNY" in line:
                return True
            else:
                return False

    # Получение статьи через поисковые слова в комментариях или контрагенте
    @staticmethod
    def get_article(comment_string, counterparty_string, is_income):
        global show_article_options

        if comment_string or counterparty_string:
            values_to_return = []

            # Поиск по комментарию
            for search_words_and_article in articles_by_search_words_for_comments.items():
                # Если ключи (tuple поисковых слов) из словаря статей совпадают с комментарием
                if any(search_word.lower() in comment_string.lower() for search_word in search_words_and_article[0]):
                    values_to_return.append(
                        search_words_and_article[1])  # Возвращаем значение (статьи) из словаря статей

            # Поиск по контрагенту
            for search_words_and_article in articles_by_search_words_for_counterpartys.items():
                if any(search_word.lower() in counterparty_string.lower() for search_word in
                       search_words_and_article[0]):
                    values_to_return.append(search_words_and_article[1])

            # Поиск по комментарию и контрагенту
            for search_words_and_article in articles_by_search_words_for_comments_or_counterpartys.items():
                if any(search_word.lower() in (" ".join([comment_string.lower(), counterparty_string.lower()]))
                       for search_word in search_words_and_article[0]):
                    values_to_return.append(search_words_and_article[1])

            # Поиск для случаев, если статья зависит от притока/оттока ДС
            if is_income:
                for search_words_and_article in articles_by_search_words_for_comments_if_income.items():
                    if any(search_word.lower() in comment_string.lower() for search_word in search_words_and_article[0]):
                        values_to_return.append(search_words_and_article[1])
                for search_words_and_article in articles_by_search_words_for_comments_or_counterpartys_if_income.items():
                    if any(search_word.lower() in (" ".join([comment_string.lower(), counterparty_string.lower()]))
                           for search_word in search_words_and_article[0]):
                        values_to_return.append(search_words_and_article[1])

            elif not is_income:
                for search_words_and_article in articles_by_search_words_for_comments_if_outcome.items():
                    if any(search_word.lower() in comment_string.lower() for search_word in search_words_and_article[0]):
                        values_to_return.append(search_words_and_article[1])
                for search_words_and_article in articles_by_search_words_for_comments_or_counterpartys_if_outcome.items():
                    if any(search_word.lower() in (" ".join([comment_string.lower(), counterparty_string.lower()]))
                           for search_word in search_words_and_article[0]):
                        values_to_return.append(search_words_and_article[1])

            # Возвращаем статью, которая идет раньше в порядке приоритета или возращаем статьи через "/"
            if values_to_return != []:
                # Убираем дубликаты
                values_to_return = list(set(values_to_return))
                # Если не стоит галочка "Показывать все варианты статей"
                if not show_article_options:
                    if len(values_to_return) > 1:
                        for article in priority_order_of_articles:
                            if article in values_to_return:
                                return article
                        return '/'.join(values_to_return)
                    else:
                        return '/'.join(values_to_return)
                else:
                    return '/'.join(values_to_return)
            else:
                return ""
        else:
            return ""

    # Поиск следующей свободной строки в таблице
    @staticmethod
    def next_available_row(worksheet):
        # Смотрим по второй колонке (первая пустая синяя)
        return len(worksheet.col_values(2))+1

    # Работа с гугл таблицами
    def upload_to_googledrive(self):
        app = BankStatementsApp.get_running_app()
        service_account = gspread.service_account(filename=BankStatementsApp.resource_path(service_account_filename))
        try:
            spreadsheet = service_account.open(google_sheet_name)
        except:
            # Вылезающее уведомление с ошибкой SpreadsheetNotFound
            if not self.google_sheet_name_error_snackbar:
                self.google_sheet_name_error_snackbar = Snackbar(text=f"Ошибка: неверное название таблицы или отсутствует интернет!",
                                                      font_size=app.font_size_value,
                                                      duration=3, size_hint_x=0.8,
                                                      pos_hint={"center_x": 0.5, "center_y": 0.1})
            self.google_sheet_name_error_snackbar.open()
        try:
            worksheet = spreadsheet.worksheet(worksheet_name)
        except:
            # Вылезающее уведомление с ошибкой WorksheetNotFound
            if not self.worksheet_name_error_snackbar:
                self.worksheet_name_error_snackbar = Snackbar(text=f"Ошибка: неверное название листа!",
                                                      font_size=app.font_size_value,
                                                      duration=3, size_hint_x=0.8,
                                                      pos_hint={"center_x": 0.5, "center_y": 0.1})
            self.worksheet_name_error_snackbar.open()

        if self.format == "txt":
            required_values = self.get_required_values()
            income_checking_account = self.get_income_checking_account()  # РасчСчет
            data_to_upload = self.get_data_to_upload(required_values, income_checking_account)
        elif self.format == "pdf":
            required_values, is_sberbank = self.get_required_values_from_pdf()
            data_to_upload = self.get_data_to_upload_from_pdf(required_values, is_sberbank)
        if data_to_upload:  # Проверяем не пустой ли файл
            # Проверка с выводом ошибки
            try:
                # Проверка, что в таблице хватает свободных строк
                if (self.next_available_row(worksheet)-1 + len(data_to_upload) <= worksheet.row_count):
                    worksheet.update(f"B{self.next_available_row(worksheet)}:{chr(ord('B') - 1 + len(data_to_upload[0]))}"
                                     f"{self.next_available_row(worksheet) + len(data_to_upload)}",
                                     data_to_upload, value_input_option='USER_ENTERED')
                    # chr(ord('A')-1+len(data_to_upload[0])) - буква алфавита по номеру начиная с заглавной A
                else:
                    print("Ошибка: недостаточно свободных строк в таблице")
                    # Вылезающее уведомление с ошибкой при выгрузке данных снизу
                    if not self.not_enough_rows_error_snackbar:
                        self.not_enough_rows_error_snackbar = Snackbar(text="Ошибка: недостаточно свободных строк в таблице",
                                                              font_size=app.font_size_value,
                                                              duration=3, size_hint_x=0.8,
                                                              pos_hint={"center_x": 0.5, "center_y": 0.1})
                    self.not_enough_rows_error_snackbar.open()
            except Exception as e:
                print("Ошибка: ", e)
                # Вылезающее уведомление с ошибкой при выгрузке данных снизу
                if not self.upload_error_snackbar:
                    self.upload_error_snackbar = Snackbar(text=f"Ошибка: {e}",
                                                        font_size=app.font_size_value,
                                                        duration=3, size_hint_x=0.8,
                                                        pos_hint={"center_x": 0.5, "center_y": 0.1})
                self.upload_error_snackbar.open()
        else:
            print("Файл не выбран или содержимое отсутствует.")
            # Вылезающее уведомление с ошибкой данных выбранного файла снизу
            if not self.data_error_snackbar:
                self.data_error_snackbar = Snackbar(text="Выбранный файл отсутствует или недействителен!",
                                                    font_size=app.font_size_value,
                                                    duration=3, size_hint_x=0.8,
                                                    pos_hint={"center_x": 0.5, "center_y": 0.1})
            self.data_error_snackbar.open()

    # Открыть окно настроек
    def open_settings(self):
        if not self.settings_dialog:
            self.settings_dialog = MDDialog(
                title="Настройки",
                type="custom",
                content_cls=SettingsDialogContent(),
                buttons=[MDFlatButton(text="Закрыть",
                                      on_release=lambda instance: self.settings_dialog.dismiss())
                ],
            )
        self.settings_dialog.open()


# Содержимое настроек
class SettingsDialogContent(MDRecycleView):
    # Обновление настроек при вводе данных
    def update_settings(self, id, text):
        if text != "":
            if id == "starting_directory":
                global starting_directory
                starting_directory = text
                store.put('starting_directory', starting_directory=starting_directory)
                self.ids.starting_directory.text = ""
            elif id == "google_sheet_name":
                global google_sheet_name
                google_sheet_name = text
                store.put('google_sheet_name', google_sheet_name=google_sheet_name)
                self.ids.google_sheet_name.text = ""
            elif id == "worksheet_name":
                global worksheet_name
                worksheet_name = text
                store.put('worksheet_name', worksheet_name=worksheet_name)
                self.ids.worksheet_name.text = ""

    # Сброс настроек
    def reset_settings(self):
        if store.exists('starting_directory'):
            store.delete('starting_directory')
        self.ids.starting_directory.text = ""
        global starting_directory, starting_directory_default
        starting_directory = starting_directory_default
        if store.exists('google_sheet_name'):
            store.delete('google_sheet_name')
        self.ids.google_sheet_name.text = ""
        global google_sheet_name, google_sheet_name_default
        google_sheet_name = google_sheet_name_default
        if store.exists('worksheet_name'):
            store.delete('worksheet_name')
        self.ids.worksheet_name.text = ""
        global worksheet_name, worksheet_name_default
        worksheet_name = worksheet_name_default
        if store.exists('show_article_options'):
            store.delete('show_article_options')
        self.ids.show_article_options_checkbox.active = False
        global show_article_options
        show_article_options = False
        if store.exists('is_cny_statement_manually'):
            store.delete('is_cny_statement_manually')
        self.ids.is_cny_statement_checkbox.active = False
        global is_cny_statement_manually
        is_cny_statement_manually = False

    # Установка галочки чекбокса
    def set_check(self, instance_check, text):
        global show_article_options
        global is_cny_statement_manually

        if instance_check.active:
            instance_check.active = False
        else:
            instance_check.active = True
        if text == "Показывать все варианты статей":
            show_article_options = instance_check.active
            store.put('show_article_options', show_article_options=show_article_options)

        elif text == "Выписка в юанях":
            is_cny_statement_manually = instance_check.active
            store.put('is_cny_statement_manually', is_cny_statement_manually=is_cny_statement_manually)


class BankStatementsApp(MDApp):
    # Для kv файла
    window_size = window_size
    is_cny_statement_manually = is_cny_statement_manually
    show_article_options = show_article_options

    font_size_value_int = 25  # Размер шрифта цифрой
    font_size_value = f"{font_size_value_int}sp"  # Размер шрифта
    bigger_font_size_value = f"{font_size_value_int*1.5}sp"  # Больший размер шрифта
    smaller_font_size_value = f"{font_size_value_int*0.7}sp"  # Меньший размер шрифта

    def build(self):
        # Устанавливаем название и иконку окна приложения
        self.title = 'Выгрузка банковских выписок в Google Таблицу "АРТЭЛЬ/ База данных"'
        self.icon = self.resource_path('artel_icon.png')
        # Цветовая палитра
        self.theme_cls.primary_palette = "BlueGray"
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

import smtplib
import requests
import xlrd
import xlwt
import mimetypes

from lxml import html
from itertools import starmap
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart


URL = {
    'USD': "https://www.moex.com/ru/derivatives/currency-rate.aspx?currency=USD_RUB",
    'EUR': "https://www.moex.com/ru/derivatives/currency-rate.aspx?currency=EUR_RUB"
}


class CourseParser(object):

    def __init__(self, currency_1, currency_2, filename):
        self.filename = filename
        self.currency_1 = currency_1
        self.currency_2 = currency_2
        self.titles = tuple(
            ["Дата", "Курс %s" % self.currency_1, "Изменение"] +
            ["Дата", "Курс %s" % self.currency_2, "Изменение"] +
            ['%s/%s' % (currency_2, currency_1)]
        )

    @staticmethod
    def get_raw_data_from_html(url):
        """
        Метод возвращает необработанные данные с сайта https://www.moex.com/.
        """
        response = requests.get(url).text
        content = html.fromstring(response)
        raw_data = content.xpath('//tr//td/text()')
        return raw_data

    def _get_single_currency_data(self, url):
        """
        Метод структурирует ранее полученные данные.
        """
        raw_data = self.get_raw_data_from_html(url)
        date_list = raw_data[0::3]
        days = len([i for i in date_list if i.split(".")[1] == date_list[0].split('.')[1]])
        daily_course = [float(i.replace(',', '.')) for i in raw_data[2::3]]
        daily_change = [daily_course[i] - daily_course[i + 1] for i in range(days)]
        return date_list[:days], daily_course[:days], daily_change

    @staticmethod
    def get_cell_styles():
        """
        Метод возвращает словарь стилей для использования при записи в файл.
        """
        styles = dict()

        # Bold title format
        style0 = xlwt.easyxf('align: horizontal left')
        style0.font.bold = True
        styles['s_title'] = style0

        # Date format
        style1 = xlwt.easyxf('align: horizontal left')
        style1.num_format_str = 'm/d/yy'
        styles['s_date'] = style1

        # Financial format
        style2 = xlwt.easyxf('align: horizontal left')
        style2.num_format_str = '_(* #,##0.00 _₽_);_(-* #,##0.00 _₽_;_(* "-"?? _₽_);_(@_)'
        styles['s_financial'] = style2

        # Numeric format
        style3 = xlwt.easyxf('align: horizontal left')
        style3.num_format_str = '0.00'
        styles['s_numeric'] = style3

        return styles

    def get_currencies_data(self):
        """
        Метод компонует данные в один массив для дальнейшей записи в файл.
        """
        first_currency = self._get_single_currency_data(URL[self.currency_1])
        second_currency = self._get_single_currency_data(URL[self.currency_2])

        cur_values = list(zip(first_currency[1], second_currency[1]))
        daily_ratio = list(starmap(lambda x, y: y / x, cur_values))
        currency_data = list(
            zip(
                first_currency[0],
                first_currency[1],
                first_currency[2],
                second_currency[0],
                second_currency[1],
                second_currency[2],
                daily_ratio
            )
        )
        return currency_data

    @staticmethod
    def set_auto_width(sheet, value, col_id):
        """
        Метод автоматически форматирует ширину ячеек.
        """
        c_width = sheet.col(col_id).width
        if (len(str(value) * 256)) > c_width:
            sheet.col(col_id).width = (len(str(value)) + 1) * 256

    def write_data_to_excel(self):
        """
        Метод записывает полученные данные в файл с применением ранее заданных стилей.
        """
        data = self.get_currencies_data()
        styles = self.get_cell_styles()

        wb = xlwt.Workbook()
        ws = wb.add_sheet(
            'Курсы валют %s к %s' % (
                self.currency_2,
                self.currency_1,
            ),
            cell_overwrite_ok=True
        )

        for k, v in enumerate(self.titles):
            self.set_auto_width(ws, v, k)
            ws.write(0, k, v, styles['s_title'])

        for key, values in enumerate(data, start=1):
            for col_id in range(7):
                self.set_auto_width(ws, values[col_id], col_id)
                if col_id in [0, 3]:
                    ws.write(key, col_id, values[col_id], styles['s_date'])
                elif col_id in [1, 4]:
                    ws.write(key, col_id, values[col_id], styles['s_financial'])
                else:
                    ws.write(key, col_id, values[col_id], styles['s_numeric'])
        filename_for_attaching = '%s_%s_%s.xls' % (self.filename, self.currency_1, self.currency_2)
        wb.save(filename_for_attaching)
        return filename_for_attaching


class SendCoursesByEmail(object):
    def __init__(self, sender, recipient, password, filename, subject):
        self.sender = sender
        self.recipient = recipient
        self.password = password
        self.filename = filename
        self.subject = subject

    def get_row_number_from_file(self):
        """
        Метод возвращает количество строк с данными.
        Шапка не учитывается при подсчете.
        """
        wb = xlrd.open_workbook(self.filename)
        ws = wb.sheet_by_index(0)
        return len(list(ws.get_rows())) - 1

    def get_message_text(self):
        """
        Метод возвращает текст сообщения.
        """
        row_number = self.get_row_number_from_file()
        if 11 <= row_number % 100 <= 14 or row_number % 10 > 5 or row_number % 5 == 0:
            char = ""
        elif row_number % 10 == 1:
            char = "у"
        else:
            char = "и"
        return "Файл %s содержит %s строк%s с данными." % (self.filename, row_number, char)

    def make_file_attachment(self):
        """
        Метод обрабатывает файл для прикрепления к сообщению.
        """
        file_type, _ = mimetypes.guess_type(self.filename)
        file_type = file_type.split('/')
        maintype, subtype = file_type[0], file_type[1]
        attachment = MIMEBase(maintype, subtype)
        with open(self.filename, "rb") as file:
            attachment.set_payload(file.read())
            file.close()
        encoders.encode_base64(attachment)
        attachment.add_header(
            'Content-Disposition',
            'attachment; filename=%s' % self.filename
        )
        return attachment

    def send_email(self):
        """
        Метод отправляет email сообщение.
        """
        msg = MIMEMultipart()
        msg['From'] = self.sender
        msg['To'] = self.recipient
        msg['Subject'] = self.subject
        msg.attach(MIMEText(self.get_message_text(), 'plain'))
        msg.attach(self.make_file_attachment())
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(self.sender, self.password)
        server.send_message(msg)
        server.quit()


if __name__ == '__main__':
    currencies = ['USD', 'EUR']
    f_name = CourseParser(currencies[0], currencies[1], 'courses').write_data_to_excel()

    message = SendCoursesByEmail(
        input('Введите электронный адрес отправителя:'),
        input('Введите электронный адрес получателя:'),
        input('Введите пароль приложения:'),
        f_name,
        f"Курсы валют(%s/RUB, %s/RUB, %s/%s)" % (
            currencies[0],
            currencies[1],
            currencies[1],
            currencies[0]
        )
    )
    message.send_email()




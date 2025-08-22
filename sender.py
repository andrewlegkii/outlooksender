import win32com.client as win32

class Sender:
    def send_email(self, to, cc, subject, body_text, data_df=None):
        """
        Отправка письма через Outlook.
        to, cc - строки с адресами через ';'
        subject - тема письма
        body_text - основной текст письма
        data_df - pandas DataFrame для вставки в письмо как таблица
        """
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.CC = cc
        mail.Subject = subject

        body_html = f"<p>{body_text}</p>"

        if data_df is not None and not data_df.empty:
            table_html = data_df.to_html(index=False, border=1)
            body_html += "<br>" + table_html

        mail.HTMLBody = body_html
        mail.Send()

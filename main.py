import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from excel_handler import load_data, load_templates, filter_by_tc
from mailer import send_email

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Рассылка ТК через Outlook")

        self.data_file = None
        self.templates_file = None
        self.data_df = None
        self.templates = None

        # Кнопки выбора файлов
        tk.Button(root, text="Выбрать Excel с данными", command=self.choose_data_file).pack(pady=5)
        tk.Button(root, text="Выбрать шаблоны писем", command=self.choose_templates_file).pack(pady=5)

        # Выбор ТК
        self.tc_var = tk.StringVar()
        self.tc_dropdown = ttk.Combobox(root, textvariable=self.tc_var, state="readonly")
        self.tc_dropdown.pack(pady=5)

        # Кнопка отправки
        tk.Button(root, text="Отправить письма", command=self.send_mails).pack(pady=10)

    def choose_data_file(self):
        self.data_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.data_file:
            self.data_df = load_data(self.data_file)
            messagebox.showinfo("Файл загружен", f"Данные загружены из:\n{self.data_file}")
            self.update_tc_dropdown()

    def choose_templates_file(self):
        self.templates_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.templates_file:
            self.templates = load_templates(self.templates_file)
            messagebox.showinfo("Файл загружен", f"Шаблоны загружены из:\n{self.templates_file}")
            self.update_tc_dropdown()

    def update_tc_dropdown(self):
        if self.data_df is not None:
            tc_list = sorted(self.data_df["ТК"].dropna().unique())
            self.tc_dropdown['values'] = tc_list

    def send_mails(self):
        tc_name = self.tc_var.get()
        if not tc_name:
            messagebox.showerror("Ошибка", "Выберите Транспортную Компанию")
            return
        if tc_name not in self.templates:
            messagebox.showerror("Ошибка", f"Нет шаблона для ТК: {tc_name}")
            return

        filtered_df = filter_by_tc(self.data_df, tc_name)
        if filtered_df.empty:
            messagebox.showinfo("Пусто", f"Нет данных для ТК: {tc_name}")
            return

        template = self.templates[tc_name]
        table_html = filtered_df.to_html(index=False)

        body = template["Body"].replace("{table_html}", table_html)
        send_email(template["To"], template["CC"], template["Subject"], body)

        messagebox.showinfo("Готово", f"Письмо отправлено для ТК: {tc_name}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()

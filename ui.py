import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class AppUI:
    def __init__(self, root, data_loader, sender):
        self.root = root
        self.data_loader = data_loader
        self.sender = sender

        self.contacts_file = None
        self.data_file = None

        # --- кнопки загрузки файлов ---
        tk.Button(root, text="Загрузить контакты", command=self.load_contacts).pack(pady=5)
        tk.Button(root, text="Загрузить данные", command=self.load_data).pack(pady=5)

        # --- выбор ТК ---
        tk.Label(root, text="Выберите транспортную компанию:").pack()
        self.company_var = tk.StringVar()
        self.company_combo = ttk.Combobox(root, textvariable=self.company_var, state="readonly")
        self.company_combo.pack(pady=5)

        # --- кнопка отправки ---
        tk.Button(root, text="Отправить письмо", command=self.send_email).pack(pady=10)

    def load_contacts(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return
        self.contacts_file = path
        self.data_loader.load_contacts(path)
        self.update_companies()
        messagebox.showinfo("OK", "Файл контактов загружен!")

    def load_data(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return
        self.data_file = path
        self.data_loader.load_data(path)
        messagebox.showinfo("OK", "Файл с данными загружен!")

    def update_companies(self):
        companies = self.data_loader.get_companies()
        self.company_combo["values"] = companies
        if companies:
            self.company_combo.current(0)

    def send_email(self):
        company = self.company_var.get()
        if not company:
            messagebox.showerror("Ошибка", "Выберите ТК")
            return

        contacts = self.data_loader.get_contacts(company)
        if not contacts:
            messagebox.showerror("Ошибка", f"Нет контактов для {company}")
            return

        # --- фиксированный текст письма ---
        subject = f"Контроль учета деревянных поддонов при поставках в Тандер {company}"
        body_text = (
            "Добрый день!\n\n"
            "Это письмо-напоминание, что мы ждем фото Товарной накладной на поддоны, "
            "сделанное водителем сразу после выгрузки продукции на РЦ Тандера."
        )

        # --- таблица для вставки ---
        df = self.data_loader.get_data_for_company(company)

        # --- отправка письма ---
        self.sender.send_email(
            to=contacts["to"],
            cc=contacts["cc"],
            subject=subject,
            body_text=body_text,
            data_df=df
        )

        messagebox.showinfo("Успех", f"Письмо для {company} отправлено!")

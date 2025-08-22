import pandas as pd

class DataLoader:
    def __init__(self):
        self.contacts_df = None
        self.data_df = None

    def load_contacts(self, file_path):
        """Загрузка файла контактов"""
        self.contacts_df = pd.read_excel(file_path)
        return self.contacts_df

    def load_data(self, file_path):
        """Загрузка файла с данными"""
        self.data_df = pd.read_excel(file_path)
        return self.data_df

    def get_companies(self):
        """Получаем список уникальных ТК из contacts.xlsx"""
        if self.contacts_df is None:
            return []
        return sorted(self.contacts_df["ТК"].dropna().unique().tolist())

    def get_contacts(self, company):
        """Получаем Email и CC по выбранной компании"""
        if self.contacts_df is None:
            return None
        row = self.contacts_df[self.contacts_df["ТК"] == company]
        if row.empty:
            return None

        to_emails = [e.strip() for e in str(row.iloc[0].get("Email", "")).split(";") if e.strip()]
        cc_emails = [e.strip() for e in str(row.iloc[0].get("Копия", "")).split(";") if e.strip()]

        return {"to": ";".join(to_emails), "cc": ";".join(cc_emails)}

    def get_data_for_company(self, company):
        """Фильтруем строки из data.xlsx по ТК"""
        if self.data_df is None:
            return None
        filtered = self.data_df[self.data_df["ТК"] == company]
        return filtered

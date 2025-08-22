import tkinter as tk
from data_loader import DataLoader
from sender import Sender
from ui import AppUI


def main():
    root = tk.Tk()
    root.title("Рассылка ТК через Outlook")

    # Создаём объекты загрузчика данных и отправителя
    loader = DataLoader()
    sender = Sender()  # Отправка через уже запущенный Outlook

    # Создаём GUI
    app = AppUI(root, loader, sender)

    root.mainloop()


if __name__ == "__main__":
    main()

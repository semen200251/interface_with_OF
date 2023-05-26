import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import threading
import os
import time
import readOF
import config_for_interface

def get_folder_click():
    """Функция описывает логику работу кнопки выбора директории для выгрузки ОФ"""
    config_for_interface.path_to_folder = filedialog.askdirectory()
    print("Selected folder:", config_for_interface.path_to_folder)
    if config_for_interface.path_to_project:
        start_button.pack(pady=10)

def get_project_click():
    """Функция описывает логику работу кнопки выбора файлов projects для выгрузки в ОФ"""
    config_for_interface.path_to_project = filedialog.askopenfilenames()
    print("Selected file:", config_for_interface.path_to_project)
    if config_for_interface.path_to_folder:
        start_button.pack(pady=10)

def update_progress(value):
    """Функция обновляет значение"""
    progress['value'] = value
    percent_label.configure(text=f"Загрузка: {value}%")

def simulate_progress():
    for path in config_for_interface.path_to_project:
        size_in_bytes = os.path.getsize(path)
        size_in_kilobytes = size_in_bytes / 1024
        timer = size_in_kilobytes/2000 - 1
        if timer <= 0:
            timer = 1
        for i in range(101):
            update_progress(i)
            time.sleep(timer/100)  # Задержка 0.1 секунды между обновлениями
        progress.stop()  # Остановка анимации загрузки

def start_progress():
    progress.pack(pady=10)
    percent_label.pack()
    threading.Thread(target=simulate_progress).start()
    threading.Thread(target=readOF.main, args=(config_for_interface.path_to_project, config_for_interface.path_to_folder)).start()
    print(config_for_interface.path_to_excel)

if __name__ == '__main__':
    folderpath = None
    filepath = None
    path_to_excel = None
    window = tk.Tk()
    window.title("Выгрузка ОФ")
    window.geometry("400x300")
    button_style = {'background': '#4CAF50', 'foreground': 'white', 'font': ('Arial', 12)}
    label1 = tk.Label(window, text="Выберите папку, в которую выгрузить ОФ", font=('Arial', 14))
    label1.pack(pady=10)
    button1 = tk.Button(window, text="Выбрать папку", command=get_folder_click, **button_style)
    button1.pack()
    label2 = tk.Label(window, text="Выберите project", font=('Arial', 14))
    label2.pack(pady=10)
    button2 = tk.Button(window, text="Выбрать файл", command=get_project_click, **button_style)
    button2.pack()
    progress = ttk.Progressbar(window, length=300, mode='determinate')
    percent_label = tk.Label(window, text="Загрузка: 0%", font=('Arial', 12))
    start_button = tk.Button(window, text="Начать загрузку", command=start_progress, **button_style)
    window.mainloop()


import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import threading
import os
import time
import readOF
import config_for_interface
import database

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
    """Функция обновляет значение количества загруженных файлов"""
    percent_label.configure(text=f"Загружено: {value} файлов")

def switch_info_labels(value):
    """Функция выводит текстовую информацию о результате загрузки"""
    succes = sum(1 for item in config_for_interface.path_to_excel if item is not None)
    info_label.configure(text=f"Загрузка завершена. Загружено: {value} файлов.\n Успешно: {succes}")

def start_progress():
    """Функция выполняет выгрузку выбранных project в ОФ"""
    percent_label.pack()
    info_label.pack()
    value = 0
    window.update()
    for path in config_for_interface.path_to_project:
        config_for_interface.path_to_excel.append(readOF.main(path, config_for_interface.path_to_folder))
        value = value + 1
        update_progress(value)
        window.update()
    percent_label.destroy()
    switch_info_labels(value)
    database.fill_data(config_for_interface.path_to_excel, config_for_interface.path_to_project, config_for_interface.path_to_folder)
    database.view_data()



if __name__ == '__main__':
    window = tk.Tk()
    database.create_database()
    window.title("Выгрузка ОФ")
    window.geometry("900x600")
    button_style = {'background': '#4CAF50', 'foreground': 'white', 'font': ('Arial', 12)}
    label1 = tk.Label(window, text="Выберите папку, в которую выгрузить ОФ", font=('Arial', 14))
    label1.pack(pady=10)
    button1 = tk.Button(window, text="Выбрать папку", command=get_folder_click, **button_style)
    button1.pack()
    label2 = tk.Label(window, text="Выберите project", font=('Arial', 14))
    label2.pack(pady=10)
    button2 = tk.Button(window, text="Выбрать файл", command=get_project_click, **button_style)
    button2.pack()
    percent_label = tk.Label(window, text="Загружено: 0 файлов", font=('Arial', 12))
    info_label = tk.Label(window, text="Пожалуйста, ожидайте, выгрузка ОФ может занимать длительное время", font=('Arial', 12))
    start_button = tk.Button(window, text="Начать загрузку", command=start_progress, **button_style)
    window.mainloop()


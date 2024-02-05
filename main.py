import tkinter as tk
import pandas as pd
from tkinter import ttk
import ast
import os

class DatabaseApplication:
    def __init__(self, root):
        """Инициализация приложения"""
        self.root = root
        self.root.title("База данных")
        self.root.geometry("580x300")
        self.root.configure(background='light grey')

        self.file_path = "articles.txt"
        self.articles = self.read_articles()  # Чтение статей из файла

        self.create_widgets()  # Создание виджетов

    def read_articles(self):
        """Чтение статей из файла и сортировка по номеру"""
        try:
            with open(self.file_path, 'r') as file:
                data = file.read()
                articles = ast.literal_eval(data)  # Преобразуем строку с данными обратно в кортеж
                articles.sort(key=lambda x: x[0])  # Сортировка статей по номеру
                return articles
        except FileNotFoundError:
            return []

    def find_closest_larger_article(self, number):
        """Поиск ближайшей статьи с большим номером"""
        for article in self.articles:
            if article[0] > number:
                return article
        return None

    def write_articles(self):
        """Запись статей в файл"""
        with open(self.file_path, 'w') as file:
            file.write(str(self.articles))  # Записываем данные в файл как строку

    def search_articles(self):
        """Поиск статей"""
        search_query = self.search_entry.get()

        self.tree.delete(*self.tree.get_children())  # Очистить отображение статей
        for article in self.articles:
            if str(search_query) == str(article[0]):
                self.tree.insert("", "end", values=article)

    def find_closest_larger_article_index(self, number):
        """Поиск индекса ближайшей статьи с большим номером"""
        for index, article in enumerate(self.articles):
            if article[0] > number:
                return index
        return len(self.articles)  # Возвращаем индекс, на котором нужно вставить новую статью

    def save_article(self, number_entry, parent_entry, left_entry, right_entry, add_window):
        """Сохранение статьи и обновление порядка"""
        number = int(number_entry.get())
        parent = int(parent_entry.get()) if parent_entry.get() else None
        left = int(left_entry.get()) if left_entry.get() else None
        right = int(right_entry.get()) if right_entry.get() else None

        article = (number, parent, left, right)
        index = self.find_closest_larger_article_index(number)
        self.articles.insert(index, article)  # Вставляем статью на подходящий индекс
        self.update_display()  # Обновляем отображение
        self.write_articles()  # Записываем изменения в файл
        add_window.destroy()

    def update_display(self):
        """Обновление отображения статей"""
        self.tree.delete(*self.tree.get_children())  # Очищаем текущее отображение
        for article in self.articles:
            self.tree.insert("", "end", values=article)

    def update_search_results(self):
        """Обновление результатов поиска при изменении текста в поле поиска"""
        search_query = self.search_entry.get()

        if not search_query:
            self.update_display()
        else:
            self.search_articles()

    def add_article(self):
        """Добавление статьи"""
        add_window = tk.Toplevel(self.root)
        add_window.title("Добавить статью")
        add_window.geometry("400x200")

        number_label = tk.Label(add_window, text="Номер статьи:")
        number_label.pack()

        number_entry = tk.Entry(add_window)
        number_entry.pack()

        parent_label = tk.Label(add_window, text="Статья-родитель:")
        parent_label.pack()

        parent_entry = tk.Entry(add_window)
        parent_entry.pack()

        left_label = tk.Label(add_window, text="Статья-левый потомок:")
        left_label.pack()

        left_entry = tk.Entry(add_window)
        left_entry.pack()

        right_label = tk.Label(add_window, text="Статья-правый потомок:")
        right_label.pack()

        right_entry = tk.Entry(add_window)
        right_entry.pack()

        save_button = tk.Button(add_window, text="Сохранить",
                                command=lambda: self.save_article(number_entry, parent_entry, left_entry, right_entry,
                                                                  add_window))
        save_button.pack()

    def delete_article(self):
        """Удаление выбранной статьи"""
        selected_item = self.tree.selection()
        if selected_item:
            index = int(selected_item[0][1:]) - 1  # Получаем индекс выделенной статьи в массиве
            if index < len(self.articles):
                del self.articles[index]  # Удаляем статью из массива
                self.tree.delete(selected_item)  # Удаляем статью из treeview
                self.write_articles()  # Сохраняем изменения в файле
            else:
                print("Индекс находится вне диапазона списка статей")

    def delete_all_articles(self):
        """Удаление всех статей"""
        self.articles = []  # Очистить список статей
        self.tree.delete(*self.tree.get_children())  # Очистить treeview
        self.write_articles()  # Сохранить изменения в файле

    def export_to_excel(self):
        """Экспорт в Excel"""
        df = pd.DataFrame(self.articles, columns=["Number", "Parent", "Left", "Right"])
        file_path = "database.xlsx"
        df.to_excel(file_path, index=False)
        os.startfile(file_path)  # Открытие excel-файла при нажатии кнопки

    def create_widgets(self):
        """Создание виджетов для отображения статей и управления базой данных"""
        columns = ("number", "parent", "left", "right")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
        self.tree.pack(side="bottom", fill="both", expand=True)

        for column in columns:
            self.tree.column(column, width=50)
            self.tree.heading(column, text=column)

        for article in self.articles:
            self.tree.insert("", "end", values=article)

        self.search_entry = tk.Entry(self.root, bg='white', fg='black', width=20)
        self.search_entry.pack(side="left", fill="x")
        self.search_entry.bind('<KeyRelease>', lambda event: self.update_search_results())

        self.search_button = tk.Button(self.root, text="Поиск", command=self.search_articles, font=('Times New Roman', 10))
        self.search_button.pack(side='left')

        self.add_button = tk.Button(self.root, text="Добавить статью", command=self.add_article, font=('Times New Roman', 10))
        self.add_button.pack(side='left')

        self.delete_button = tk.Button(self.root, text="Удалить статью", command=self.delete_article, font=('Times New Roman', 10))
        self.delete_button.pack(side='left')

        self.delete_all_button = tk.Button(self.root, text="Удалить все статьи", command=self.delete_all_articles, font=('Times New Roman', 10))
        self.delete_all_button.pack(side='left')

        self.export_button = tk.Button(self.root, text="Export to Excel", bg='green', command=self.export_to_excel, font=('Times New Roman', 10))
        self.export_button.pack(side='left')

        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Times New Roman', 12))
        style.configure("Treeview", font=('Times New Roman', 11))

        frame = tk.Frame(self.root)
        frame.pack(pady=10)
        self.tree.pack(fill="both", expand=True, padx=0)
        self.search_entry.pack(pady=10)
        self.search_button.pack(side="left")
        self.add_button.pack(side="left")
        self.delete_button.pack(side="left")
        self.delete_all_button.pack(side="left")

if __name__ == "__main__":
    root = tk.Tk()
    app = DatabaseApplication(root)
    root.mainloop()
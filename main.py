import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import glob
import win32print
import win32api
import win32con
import time

class ModernBatchPrinter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Потоковая печать")
        self.root.geometry("640x850")
        self.root.configure(bg='#f5f5f5')
        
        # Устанавливаем иконку (если есть)
        try:
            self.root.iconbitmap('printer.ico')
        except:
            pass
        
        self.folders = []
        self.setup_modern_ui()
        
    def setup_modern_ui(self):
        # Основной контейнер
        main_frame = tk.Frame(self.root, bg='#f5f5f5')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Заголовок
        header_frame = tk.Frame(main_frame, bg='#f5f5f5')
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(header_frame, text="🖨️Потоковая печать", 
                font=('Segoe UI', 18, 'bold'),
                bg='#f5f5f5', fg='#2c3e50').pack(side=tk.LEFT)
        
        # Карточка с папками
        folder_card = tk.Frame(main_frame, bg='white', relief=tk.RIDGE, bd=1)
        folder_card.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Заголовок карточки
        tk.Label(folder_card, text="Выбранные папки", 
                font=('Segoe UI', 11, 'bold'),
                bg='white', fg='#34495e').pack(anchor=tk.W, padx=15, pady=10)
        
        # Список папок с прокруткой
        list_frame = tk.Frame(folder_card, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.folder_listbox = tk.Listbox(list_frame, 
                                        font=('Segoe UI', 10),
                                        bg='white',
                                        relief=tk.FLAT,
                                        selectbackground='#3498db',
                                        selectforeground='white',
                                        yscrollcommand=scrollbar.set)
        self.folder_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.folder_listbox.yview)
        
        # Кнопки управления папками
        btn_frame = tk.Frame(folder_card, bg='white')
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        
        add_btn = tk.Button(btn_frame, text="📁 Добавить папку",
                          command=self.add_folder,
                          font=('Segoe UI', 10),
                          bg='#3498db',
                          fg='white',
                          relief=tk.FLAT,
                          padx=20,
                          pady=8,
                          cursor='hand2')
        add_btn.pack(side=tk.LEFT)
        
        remove_btn = tk.Button(btn_frame, text="🗑️ Удалить",
                             command=self.remove_folder,
                             font=('Segoe UI', 10),
                             bg='#e74c3c',
                             fg='white',
                             relief=tk.FLAT,
                             padx=20,
                             pady=8,
                             cursor='hand2')
        remove_btn.pack(side=tk.LEFT, padx=10)
        
        # Карточка настроек
        settings_card = tk.Frame(main_frame, bg='white', relief=tk.RIDGE, bd=1)
        settings_card.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(settings_card, text="Настройки печати",
                font=('Segoe UI', 11, 'bold'),
                bg='white', fg='#34495e').pack(anchor=tk.W, padx=15, pady=10)
        
        # Грид для настроек
        settings_grid = tk.Frame(settings_card, bg='white')
        settings_grid.pack(fill=tk.X, padx=15, pady=(0, 15))
        
        # Принтер
        tk.Label(settings_grid, text="Принтер:", 
                font=('Segoe UI', 10),
                bg='white').grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.printer_var = tk.StringVar()
        printers = [win32print.GetDefaultPrinter()]
        try:
            printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
        except:
            pass
        
        printer_combo = ttk.Combobox(settings_grid, 
                                    textvariable=self.printer_var,
                                    values=printers,
                                    font=('Segoe UI', 10),
                                    state='readonly',
                                    width=50)
        printer_combo.grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)
        printer_combo.set(printers[0] if printers else "")
        
        # Типы файлов
        tk.Label(settings_grid, text="Типы файлов:", 
                font=('Segoe UI', 10),
                bg='white').grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.file_types_var = tk.StringVar(value="*.xls, *.doc")
        file_entry = tk.Entry(settings_grid, 
                             textvariable=self.file_types_var,
                             font=('Segoe UI', 10),
                             relief=tk.SOLID,
                             width=53,
                             bd=1)
        file_entry.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Двусторонняя печать
        self.duplex_var = tk.BooleanVar(value=False)
        duplex_check = tk.Checkbutton(settings_grid, 
                                     text="Двусторонняя печать",
                                     variable=self.duplex_var,
                                     font=('Segoe UI', 10),
                                     bg='white',
                                     activebackground='white',
                                     cursor='hand2')
        duplex_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Ориентация страницы
        tk.Label(settings_grid, text="Ориентация:", 
                font=('Segoe UI', 10),
                bg='white').grid(row=3, column=0, sticky=tk.W, pady=5)
        
        self.orientation_var = tk.StringVar(value="Книжная")
        orientation_combo = ttk.Combobox(settings_grid,
                                        textvariable=self.orientation_var,
                                        values=["Книжная", "Альбомная"],
                                        font=('Segoe UI', 10),
                                        state='readonly',
                                        width=20)
        orientation_combo.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Большая кнопка печати
        print_btn = tk.Button(main_frame, 
                            text="🚀 НАЧАТЬ ПЕЧАТЬ",
                            command=self.start_printing,
                            font=('Segoe UI', 12, 'bold'),
                            bg='#2ecc71',
                            fg='white',
                            relief=tk.FLAT,
                            padx=40,
                            pady=15,
                            cursor='hand2')
        print_btn.pack(pady=20)
        
        # Эффекты при наведении
        for btn in [add_btn, remove_btn, print_btn]:
            self.add_hover_effect(btn)
            
        # Статус бар
        self.status_bar = tk.Label(self.root, 
                                  text="Готов к работе",
                                  bg='#34495e',
                                  fg='white',
                                  font=('Segoe UI', 9),
                                  anchor=tk.W,
                                  padx=10)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def add_hover_effect(self, button):
        original_color = button.cget('background')
        
        def on_enter(e):
            if button.cget('state') != 'disabled':
                # Делаем цвет темнее
                colors = {
                    '#3498db': '#2980b9',  # Синий
                    '#e74c3c': '#c0392b',  # Красный
                    '#2ecc71': '#27ae60'   # Зеленый
                }
                new_color = colors.get(original_color, original_color)
                button.config(bg=new_color)
        
        def on_leave(e):
            button.config(bg=original_color)
        
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
    
    def add_folder(self):
        folder = filedialog.askdirectory()
        if folder and folder not in self.folders:
            self.folders.append(folder)
            self.folder_listbox.insert(tk.END, folder)
            self.status_bar.config(text=f"Добавлена папка: {folder}")
    
    def remove_folder(self):
        selection = self.folder_listbox.curselection()
        if selection:
            index = selection[0]
            folder = self.folders[index]
            self.folder_listbox.delete(index)
            del self.folders[index]
            self.status_bar.config(text=f"Удалена папка: {folder}")
    
    def start_printing(self):
        if not self.folders:
            messagebox.showerror("Ошибка", "Выберите хотя бы одну папку")
            return
        
        self.status_bar.config(text="Идет печать...")
        self.root.update()  # Обновляем интерфейс
        
        file_patterns = [pattern.strip() for pattern in self.file_types_var.get().split(",")]
        total_files = 0
        
        # Получаем имя принтера
        printer_name = self.printer_var.get()
        if not printer_name:
            messagebox.showerror("Ошибка", "Не выбран принтер")
            return
        
        # Получаем настройки
        duplex = self.duplex_var.get()
        orientation = self.orientation_var.get()
        
        # Настройка принтера один раз перед печатью всех файлов
        try:
            self.configure_printer(printer_name, duplex, orientation)
        except Exception as e:
            messagebox.showwarning("Внимание", f"Не удалось применить настройки принтера:\n{str(e)}\nПродолжаем печать с настройками по умолчанию.")
        
        for folder in self.folders:
            for pattern in file_patterns:
                files = glob.glob(os.path.join(folder, pattern))
                for file_path in files:
                    try:
                        # Отправляем файл на печать
                        win32api.ShellExecute(
                            0,
                            "printto",
                            file_path,
                            f'"{printer_name}"',
                            ".",
                            0
                        )
                        
                        total_files += 1
                        file_name = os.path.basename(file_path)
                        self.status_bar.config(text=f"Отправлено на печать: {file_name}")
                        self.root.update()  # Обновляем интерфейс
                        
                        # Небольшая задержка между файлами
                        time.sleep(1)
                        
                    except Exception as e:
                        print(f"Ошибка: {file_path} - {e}")
                        messagebox.showerror("Ошибка печати", f"Файл: {os.path.basename(file_path)}\nОшибка: {str(e)}")
        
        messagebox.showinfo("Готово", f"Отправлено {total_files} файлов на печать")
        self.status_bar.config(text=f"Готово. Отправлено {total_files} файлов")
    
    def configure_printer(self, printer_name, duplex, orientation):
        """
        Настраивает параметры принтера перед печатью
        """
        try:
            # Открываем принтер
            printer_handle = win32print.OpenPrinter(printer_name)
            
            # Получаем текущие настройки принтера
            defaults = win32print.GetPrinter(printer_handle, 2)
            devmode = defaults["pDevMode"]
            
            # Настройка двусторонней печати
            if duplex:
                # 1 = горизонтальная двусторонняя, 2 = вертикальная двусторонняя
                devmode.Duplex = 2  # Вертикальная двусторонняя печать
            else:
                devmode.Duplex = 1  # Односторонняя печать
            
            # Настройка ориентации
            if orientation == "Альбомная":
                devmode.Orientation = win32con.DMORIENT_LANDSCAPE  # Альбомная
            else:
                devmode.Orientation = win32con.DMORIENT_PORTRAIT   # Книжная
            
            # Сохраняем изменения
            defaults["pDevMode"] = devmode
            win32print.SetPrinter(printer_handle, 2, defaults, 0)
            
            # Закрываем принтер
            win32print.ClosePrinter(printer_handle)
            
            print(f"Принтер настроен: Двусторонняя={duplex}, Ориентация={orientation}")
            
        except Exception as e:
            print(f"Ошибка настройки принтера: {e}")
            raise
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ModernBatchPrinter()
    app.run()
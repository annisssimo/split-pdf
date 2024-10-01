import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from PIL import Image
import fitz  # PyMuPDF

# Максимальный размер файла в байтах (5 МБ)
MAX_SIZE = 5 * 1024 * 1024

CONFIG_FILE = "config.txt"

# Функция для загрузки пути из файла конфигурации
def load_previous_paths():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            paths = f.readlines()
            return paths[0].strip(), paths[1].strip()
    else:
        return None, None

# Функция для сохранения выбранных путей
def save_paths(input_dir, output_dir):
    with open(CONFIG_FILE, "w") as f:
        f.write(f"{input_dir}\n")
        f.write(f"{output_dir}\n")

# Функции сжатия и обработки PDF/изображений
def compress_images_in_pdf(pdf_path, output_path, log_text):
    doc = fitz.open(pdf_path)
    for page in doc:
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.save(output_path, format="PDF", quality=80)
    doc.close()
    log_message(log_text, f"Сжатие изображений в PDF {pdf_path} завершено.\n", "info")

def compress_pdf(pdf_path, output_path, target_size, log_text):
    temp_output_path = output_path + "_compressed.pdf"
    compress_images_in_pdf(pdf_path, temp_output_path, log_text)
    
    if os.path.getsize(temp_output_path) > target_size:
        log_message(log_text, f"Ошибка: не удалось сжать файл {output_path} до необходимого размера.\n", "error")
    else:
        os.replace(temp_output_path, output_path)
        log_message(log_text, f"Файл {output_path} успешно сжат до нужного размера.\n", "info")

def compress_image(image_path, output_path, target_size, log_text):
    with Image.open(image_path) as img:
        quality = 95
        while os.path.getsize(output_path) > target_size and quality > 10:
            img.save(output_path, format="JPEG", quality=quality)
            quality -= 5
        log_message(log_text, f"Изображение {image_path} сжато до {quality}% качества.\n", "info")

def split_pdf(file_path, output_dir, log_text):
    pdf_document = fitz.open(file_path)
    base_file_name = os.path.splitext(os.path.basename(file_path))[0]
    log_message(log_text, f"Разделение PDF {file_path} на отдельные страницы...\n", "info")
    
    for page_num in range(len(pdf_document)):
        pdf_writer = fitz.open()
        pdf_writer.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
        page_pdf_path = os.path.join(output_dir, f"{base_file_name}({page_num + 1}).pdf")
        pdf_writer.save(page_pdf_path)
        pdf_writer.close()
        log_message(log_text, f"Страница {page_num + 1} сохранена как отдельный PDF: {page_pdf_path}\n", "success")
        
        if os.path.getsize(page_pdf_path) > MAX_SIZE:
            log_message(log_text, f"Файл {page_pdf_path} превышает 5 Мбайт, выполняем сжатие...\n", "warning")
            compress_pdf(page_pdf_path, page_pdf_path, MAX_SIZE, log_text)

def convert_image_to_pdf_or_tiff(image_path, output_dir, output_format, log_text):
    base_file_name = os.path.splitext(os.path.basename(image_path))[0]
    log_message(log_text, f"Конвертация изображения {image_path} в формат {output_format}...\n", "info")

    if output_format.upper() == "TIFF":
        tiff_path = os.path.join(output_dir, f"{base_file_name}.tiff")
        with Image.open(image_path) as img:
            img.save(tiff_path, format="TIFF")
        log_message(log_text, f"Изображение сохранено как TIFF: {tiff_path}\n", "success")

        if os.path.getsize(tiff_path) > MAX_SIZE:
            log_message(log_text, f"Файл {tiff_path} превышает 5 Мбайт, выполняем сжатие...\n", "warning")
            compress_image(image_path, tiff_path, MAX_SIZE, log_text)

    elif output_format.upper() == "PDF":
        pdf_path = os.path.join(output_dir, f"{base_file_name}.pdf")
        with Image.open(image_path) as img:
            img.save(pdf_path, format="PDF")
        log_message(log_text, f"Изображение сохранено как PDF: {pdf_path}\n", "success")

        if os.path.getsize(pdf_path) > MAX_SIZE:
            log_message(log_text, f"Файл {pdf_path} превышает 5 Мбайт, выполняем сжатие...\n", "warning")
            compress_pdf(pdf_path, pdf_path, MAX_SIZE, log_text)

# Функция для обработки файлов
def process_files(input_path, output_dir, image_output_format, log_text):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for file_name in os.listdir(input_path):
        file_path = os.path.join(input_path, file_name)
        if file_name.endswith(".pdf"):
            log_message(log_text, f"Обрабатываем PDF: {file_name}\n", "info")
            split_pdf(file_path, output_dir, log_text)
        elif file_name.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff')):
            log_message(log_text, f"Обрабатываем изображение: {file_name}\n", "info")
            convert_image_to_pdf_or_tiff(file_path, output_dir, image_output_format, log_text)
        else:
            log_message(log_text, f"Пропущен файл {file_name}: неподдерживаемый формат.\n", "error")

# Функция для вывода сообщений с цветовым тегом
def log_message(log_text, message, msg_type):
    if msg_type == "info":
        log_text.insert(tk.END, message, "info")
    elif msg_type == "warning":
        log_text.insert(tk.END, message, "warning")
    elif msg_type == "error":
        log_text.insert(tk.END, message, "error")
    elif msg_type == "success":
        log_text.insert(tk.END, message, "success")
    log_text.see(tk.END)  # Прокрутка к последнему сообщению

# Интерфейс с использованием Tkinter
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF и изображение сжатие")
        self.geometry("600x500")

        # Загружаем предыдущие пути
        input_dir, output_dir = load_previous_paths()

        # Если пути не загружены, используем пути по умолчанию
        self.default_input_dir = input_dir if input_dir else "/Users/annisssimo/Downloads/files"
        self.default_output_dir = output_dir if output_dir else "/Users/annisssimo/Downloads/pages"

        # Поле для выбора входной директории
        self.input_dir_label = tk.Label(self, text=f"Входная директория: {self.default_input_dir}")
        self.input_dir_label.pack(pady=5)
        self.input_dir_button = tk.Button(self, text="Изменить", command=self.select_input_dir)
        self.input_dir_button.pack(pady=5)

        # Поле для выбора выходной директории
        self.output_dir_label = tk.Label(self, text=f"Выходная директория: {self.default_output_dir}")
        self.output_dir_label.pack(pady=5)
        self.output_dir_button = tk.Button(self, text="Изменить", command=self.select_output_dir)
        self.output_dir_button.pack(pady=5)

        # Поле для выбора формата вывода изображений
        self.format_label = tk.Label(self, text="Формат вывода изображений:")
        self.format_label.pack(pady=5)
        self.format_var = tk.StringVar(value="PDF")
        self.format_option = tk.OptionMenu(self, self.format_var, "PDF", "TIFF")
        self.format_option.pack(pady=5)

        # Кнопка для запуска процесса обработки
        self.process_button = tk.Button(self, text="Запустить обработку", command=self.process_files)
        self.process_button.pack(pady=20)

        # Поле для вывода логов с цветными сообщениями
        self.log_text = scrolledtext.ScrolledText(self, width=70, height=15)
        self.log_text.pack(pady=5)

        # Настраиваем цветовые теги для разных типов сообщений
        self.log_text.tag_configure("info", foreground="blue")
        self.log_text.tag_configure("warning", foreground="orange")
        self.log_text.tag_configure("error", foreground="red")
        self.log_text.tag_configure("success", foreground="green")

        # Поля для путей
        self.input_dir = self.default_input_dir
        self.output_dir = self.default_output_dir

    def select_input_dir(self):
        selected_dir = filedialog.askdirectory(initialdir=self.default_input_dir)
        if selected_dir:
            self.input_dir = selected_dir
            self.input_dir_label.config(text=f"Входная директория: {self.input_dir}")
            save_paths(self.input_dir, self.output_dir)

    def select_output_dir(self):
        selected_dir = filedialog.askdirectory(initialdir=self.default_output_dir)
        if selected_dir:
            self.output_dir = selected_dir
            self.output_dir_label.config(text=f"Выходная директория: {self.output_dir}")
            save_paths(self.input_dir, self.output_dir)

    def process_files(self):
        if self.input_dir and self.output_dir:
            log_message(self.log_text, "Начинаем обработку файлов...\n", "info")
            process_files(self.input_dir, self.output_dir, self.format_var.get(), self.log_text)
            log_message(self.log_text, "Обработка завершена!\n", "success")
        else:
            messagebox.showwarning("Ошибка", "Выберите входную и выходную директории.")

# Запуск приложения
if __name__ == "__main__":
    app = App()
    app.mainloop()

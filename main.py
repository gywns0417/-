import tkinter as tk
from tkinter import simpledialog
from tkinter import messagebox
import pandas as pd
from datetime import datetime
import pickle
import os
import subprocess
import sys

class OrderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("소영 전용 주문 체크 프로그램")

        self.save_dir = self.get_save_directory()
        self.excel_save_dir = os.path.join(self.save_dir, '저장된 엑셀 파일들')
        self.ensure_directory_exists(self.excel_save_dir)

        self.items = self.load_data()

        self.frame = tk.Frame(self.root)
        self.frame.pack()

        self.add_product_button = tk.Button(self.frame, text="상품 추가", command=self.add_product)
        self.add_product_button.grid(row=0, column=0, padx=10, pady=10)

        self.save_button = tk.Button(self.frame, text="엑셀로 저장...", command=self.save_to_excel)
        self.save_button.grid(row=0, column=1, padx=10, pady=10)

        self.open_folder_button = tk.Button(self.frame, text="저장 폴더 열기", command=self.open_save_folder)
        self.open_folder_button.grid(row=0, column=2, padx=10, pady=10)

        self.update_ui()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def get_save_directory(self):
        if hasattr(sys, 'frozen'):
            return os.path.join(os.path.dirname(sys.executable), 'save_data')
        else:
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'save_data')

    def ensure_directory_exists(self, path):
        if not os.path.exists(path):
            os.makedirs(path)

    def load_data(self):
        self.ensure_directory_exists(self.save_dir)
        data_file = os.path.join(self.save_dir, "order_data.pkl")
        if os.path.exists(data_file):
            with open(data_file, "rb") as f:
                return pickle.load(f)
        return {}

    def save_data(self):
        self.ensure_directory_exists(self.save_dir)
        data_file = os.path.join(self.save_dir, "order_data.pkl")
        with open(data_file, "wb") as f:
            pickle.dump(self.items, f)

    def on_closing(self):
        self.save_data()
        self.root.destroy()

    def add_product(self):
        product_name = simpledialog.askstring("입력", "상품명을 입력하셈~")
        if product_name:
            if product_name in self.items:
                messagebox.showwarning("경고", "이미 존재하는 상품명임~")
            else:
                self.items[product_name] = 0
                self.update_ui()

    def remove_product(self, product_name):
        if product_name in self.items:
            del self.items[product_name]
            self.update_ui()

    def increase_quantity(self, product_name):
        self.items[product_name] += 1
        self.update_ui()

    def decrease_quantity(self, product_name):
        if self.items[product_name] > 0:
            self.items[product_name] -= 1
            self.update_ui()

    def set_quantity(self, product_name):
        quantity = simpledialog.askinteger("입력", f"{product_name} 의 주문 수량을 입력하쎄엠")
        if quantity is not None:
            self.items[product_name] = quantity
            self.update_ui()

    def save_to_excel(self):
        self.ensure_directory_exists(self.excel_save_dir)
        if not self.items:
            messagebox.showwarning("경고", "아무런 상품도 추가 안 했잖슴~")
            return

        df = pd.DataFrame(list(self.items.items()), columns=['상품명', '수량'])
        current_date = datetime.now().strftime('%Y년 %m월 %d일')
        file_name = os.path.join(self.excel_save_dir, f'{current_date} 주문 현황.xlsx')
        file_name = self.get_unique_file_name(file_name)
        df.to_excel(file_name, index=False)
        messagebox.showinfo("Info", f"Saved to {file_name}")

    def get_unique_file_name(self, file_name):
        base, extension = os.path.splitext(file_name)
        counter = 1
        unique_file_name = file_name
        while os.path.exists(unique_file_name):
            unique_file_name = f"{base} ({counter}){extension}"
            counter += 1
        return unique_file_name

    def open_save_folder(self):
        if os.name == 'nt':  # Windows
            subprocess.Popen(f'explorer {self.excel_save_dir}')
        elif os.name == 'posix':  # macOS, Linux
            subprocess.Popen(['open', self.excel_save_dir])

    def update_ui(self):
        for widget in self.frame.winfo_children():
            if isinstance(widget, tk.Button) and widget not in [self.add_product_button, self.save_button, self.open_folder_button]:
                widget.destroy()
            if isinstance(widget, tk.Label):
                widget.destroy()
            if isinstance(widget, tk.Entry):
                widget.destroy()

        row = 1
        for product_name, quantity in self.items.items():
            tk.Label(self.frame, text=product_name).grid(row=row, column=0, padx=10, pady=5)
            tk.Button(self.frame, text="-", command=lambda pn=product_name: self.decrease_quantity(pn)).grid(row=row, column=1, padx=5)

            quantity_entry = tk.Entry(self.frame, width=5)
            quantity_entry.insert(0, str(quantity))
            quantity_entry.bind("<Return>", lambda e, pn=product_name: self.set_quantity_entry(pn, quantity_entry))
            quantity_entry.grid(row=row, column=2, padx=5)

            tk.Button(self.frame, text="+", command=lambda pn=product_name: self.increase_quantity(pn)).grid(row=row, column=3, padx=5)
            tk.Button(self.frame, text="상품 제거", command=lambda pn=product_name: self.remove_product(pn)).grid(row=row, column=4, padx=5)
            row += 1

    def set_quantity_entry(self, product_name, entry_widget):
        try:
            quantity = int(entry_widget.get())
            self.items[product_name] = quantity
            self.update_ui()
        except ValueError:
            messagebox.showwarning("경고", "잘못된 숫자임~ 1 이상 정수만 입력하쎔~")
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, str(self.items[product_name]))

if __name__ == "__main__":
    root = tk.Tk()
    app = OrderApp(root)
    root.mainloop()


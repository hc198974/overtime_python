import os
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("串休管理")
        self.root.geometry("480x280")
        self.root.resizable(False, False)
        self.root.eval("tk::PlaceWindow . center")

        style = ttk.Style()
        style.theme_use("vista")
        style.configure("Display.TEntry", foreground="#555")

        font = ("Microsoft YaHei", 14)
        self.root.option_add("*Font", font)

        main_frame = ttk.Frame(self.root, padding="30 25 30 12")
        main_frame.pack(fill=tk.BOTH, expand=True)

        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(3, weight=1)

        self.id_name_map = {}
        self.load_data()

        row_h = 14

        ttk.Label(main_frame, text="职号：").grid(row=0, column=0, sticky=tk.E, pady=row_h)
        self.id_var = tk.StringVar()
        self.entry_id = ttk.Entry(main_frame, textvariable=self.id_var)
        self.entry_id.grid(row=0, column=1, sticky=tk.EW, pady=row_h, padx=(6, 0))

        ttk.Label(main_frame, text="姓名：").grid(row=0, column=2, sticky=tk.E, pady=row_h, padx=(20, 0))
        self.name_var = tk.StringVar()
        self.entry_name = ttk.Entry(main_frame, state="readonly", textvariable=self.name_var)
        self.entry_name.grid(row=0, column=3, sticky=tk.EW, pady=row_h, padx=(6, 0))

        ttk.Label(main_frame, text="串休时长：").grid(row=1, column=0, sticky=tk.E, pady=row_h)
        self.entry_hours = ttk.Entry(main_frame)
        self.entry_hours.grid(row=1, column=1, columnspan=3, sticky=tk.EW, pady=row_h, padx=(6, 0))

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=(18, 4))
        self.btn_submit = ttk.Button(btn_frame, text="计算串休", style="Big.TButton")

        style.configure("Big.TButton", font=("Microsoft YaHei", 14))
        self.btn_submit.pack()

        self.id_var.trace_add("write", self.on_id_changed)
        self.entry_id.focus()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def load_data(self):
        excel_path = os.path.join(os.path.dirname(__file__), "计算结果.xlsx")
        if not os.path.exists(excel_path):
            messagebox.showerror("错误", f"找不到文件：{excel_path}")
            return
        try:
            wb = load_workbook(excel_path, read_only=True)
            if "记录表" not in wb.sheetnames:
                messagebox.showerror("错误", '找不到工作表 "记录表"')
                wb.close()
                return
            ws = wb["记录表"]
            count = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] is not None:
                    key = str(int(row[0])) if isinstance(row[0], float) else str(row[0]).strip()
                    val = str(row[1]).strip() if row[1] is not None else ""
                    self.id_name_map[key] = val
                    count += 1
            wb.close()
            self.status_text = f"已加载 {count} 条记录"
        except Exception as e:
            messagebox.showerror("错误", f"读取 Excel 失败：{e}")

    def on_id_changed(self, *_):
        id_ = self.id_var.get().strip()
        if not id_:
            self.name_var.set("")
            return
        name = self.id_name_map.get(id_, "")
        if name:
            self.name_var.set(name)
        else:
            self.name_var.set("")

    def on_close(self):
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()

import tkinter as tk
from tkinter import ttk, scrolledtext
import sys
import threading
from datetime import datetime
import os

# Import the processing functions from existing scripts
# We need to make sure the current directory is in sys.path
sys.path.append(os.getcwd())

try:
    import hy
    import cz
except ImportError as e:
    print(f"Error importing modules: {e}")
    # We will handle this in the GUI if modules are missing, but for now let's assume they are there
    pass

class PrintLogger:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

class FuelManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("燃料管理系统数据处理助手")
        self.root.geometry("800x600")
        
        # Configure style
        style = ttk.Style()
        try:
            style.theme_use('clam')  
        except tk.TclError:
            pass # Use default theme if clam is not available
        
        # Main container with padding
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        header_label = ttk.Label(
            main_frame, 
            text="燃料管理数据自动处理系统", 
            font=("Microsoft YaHei UI", 16, "bold")
        )
        header_label.pack(pady=(0, 20))

        # Buttons Frame
        btn_frame = ttk.LabelFrame(main_frame, text="操作面板", padding="15")
        btn_frame.pack(fill=tk.X, pady=(0, 20))

        # Grid layout for buttons
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)

        # Analysis Button
        self.btn_hy = ttk.Button(
            btn_frame, 
            text="① 执行化验月报汇总", 
            command=self.start_hy_task
        )
        self.btn_hy.grid(row=0, column=0, padx=10, pady=10, ipady=10, sticky="ew")

        # Weight Button
        self.btn_cz = ttk.Button(
            btn_frame, 
            text="② 执行称重月报汇总", 
            command=self.start_cz_task
        )
        self.btn_cz.grid(row=0, column=1, padx=10, pady=10, ipady=10, sticky="ew")

        # Description Label
        desc_label = ttk.Label(
            btn_frame, 
            text="说明：请确保相关Excel报表已放入对应的文件夹中（无人值守化验月报 / 无人值守称重月报）",
            font=("Microsoft YaHei UI", 9),
            foreground="#666666"
        )
        desc_label.grid(row=1, column=0, columnspan=2, pady=(10, 0))

        # Log Area
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_area = scrolledtext.ScrolledText(
            log_frame, 
            state='disabled', 
            font=("Consolas", 10),
            bg="#F0F0F0"
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)

        # Status Bar
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(
            root, 
            textvariable=self.status_var, 
            relief=tk.SUNKEN, 
            anchor=tk.W,
            padding=(5, 2)
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Redirect stdout
        sys.stdout = PrintLogger(self.log_area)
        sys.stderr = PrintLogger(self.log_area)

        self.log("系统启动完成。")

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        print(f"[{timestamp}] {message}")

    def start_hy_task(self):
        self.disable_buttons()
        self.status_var.set("正在执行化验汇总...")
        self.log(">>> 开始执行化验月报汇总任务...")
        threading.Thread(target=self.run_hy_task, daemon=True).start()

    def start_cz_task(self):
        self.disable_buttons()
        self.status_var.set("正在执行称重汇总...")
        self.log(">>> 开始执行称重月报汇总任务...")
        threading.Thread(target=self.run_cz_task, daemon=True).start()

    def run_hy_task(self):
        try:
            hy.run_analysis()
            self.log("<<< 化验汇总任务完成。")
        except Exception as e:
            self.log(f"!!! 任务出错: {e}")
        finally:
            self.root.after(0, self.enable_buttons)
            self.root.after(0, lambda: self.status_var.set("就绪"))

    def run_cz_task(self):
        try:
            cz.run_weight_processing()
            self.log("<<< 称重汇总任务完成。")
        except Exception as e:
            self.log(f"!!! 任务出错: {e}")
        finally:
            self.root.after(0, self.enable_buttons)
            self.root.after(0, lambda: self.status_var.set("就绪"))

    def disable_buttons(self):
        self.btn_hy.config(state='disabled')
        self.btn_cz.config(state='disabled')

    def enable_buttons(self):
        self.btn_hy.config(state='normal')
        self.btn_cz.config(state='normal')

if __name__ == "__main__":
    root = tk.Tk()
    # Simple setup for high DPI text rendering on Windows
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
        
    app = FuelManagementApp(root)
    root.mainloop()

import sys
import pandas as pd
import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.ticker import MaxNLocator

class SignalAnalyzerGUI:
    def __init__(self, root, data):
        self.root = root
        self.df = data
        self.root.title("Visualization Data Pro")
        self.root.geometry("1450x900")
        self.all_columns = list(self.df.columns)
        self.lines = []  
        self.setup_ui()

    def setup_ui(self):
        # --- Khung biểu đồ bên TRÁI ---
        self.chart_frame = ttk.Frame(self.root)
        self.chart_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.top_ctrl_frame = ttk.Frame(self.chart_frame)
        self.top_ctrl_frame.pack(side=tk.TOP, fill=tk.X, pady=2)

        # Tắt constrained_layout để chỉnh lề thủ công
        self.fig, self.ax = plt.subplots() 
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_frame)

        self.toolbar = NavigationToolbar2Tk(self.canvas, self.top_ctrl_frame)
        self.toolbar.set_message = lambda x: None 
        self.toolbar.update()
        self.toolbar.pack(side=tk.LEFT)

        self.tooltip_enabled = tk.BooleanVar(value=True)
        self.cb_tooltip = tk.Checkbutton(self.top_ctrl_frame, text="View Point", variable=self.tooltip_enabled)
        self.cb_tooltip.pack(side=tk.LEFT, padx=10)

        self.coord_label = tk.Label(self.top_ctrl_frame, text="")
        self.coord_label.pack(side=tk.RIGHT, padx=15)

        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # --- Khung điều khiển bên PHẢI (Sidebar) ---
        self.right_frame = ttk.Frame(self.root)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10)

        # --- ĐÃ XÓA PHẦN NHẬP TIÊU ĐỀ Ở ĐÂY ---

        # 1. Chọn trục X (Đã đánh số lại)
        tk.Label(self.right_frame, text="1. Chọn Trục X:", font=('Arial', 9, 'bold')).pack(anchor="w", pady=(0,0))
        self.x_axis_var = tk.StringVar()
        self.x_combo = ttk.Combobox(self.right_frame, textvariable=self.x_axis_var, state="readonly")
        self.x_combo['values'] = self.all_columns
        if self.all_columns: self.x_combo.current(0)
        self.x_combo.pack(fill=tk.X, pady=5)
        self.x_combo.bind("<<ComboboxSelected>>", lambda e: self.update_chart())

        # 1. Tạo một nhãn tiêu đề với định dạng in đậm (bold)
        label_title = tk.Label(self.right_frame, text=" 2. Chỉnh cận trục (Limits) ", font=('Arial', 9, 'bold'))

        # 2. Gắn nhãn này vào LabelFrame thay vì dùng thuộc tính text
        limit_frame = ttk.LabelFrame(self.right_frame, labelwidget=label_title)
        limit_frame.pack(fill=tk.X, pady=10, ipady=5)

        # X-Limits
        tk.Label(limit_frame, text="X-min:").grid(row=0, column=0, padx=2)
        self.xmin_var = tk.StringVar()
        ttk.Entry(limit_frame, textvariable=self.xmin_var, width=8).grid(row=0, column=1)
        
        tk.Label(limit_frame, text="X-max:").grid(row=0, column=2, padx=2)
        self.xmax_var = tk.StringVar()
        ttk.Entry(limit_frame, textvariable=self.xmax_var, width=8).grid(row=0, column=3)

        # Y-Limits
        tk.Label(limit_frame, text="Y-min:").grid(row=1, column=0, padx=2, pady=5)
        self.ymin_var = tk.StringVar()
        ttk.Entry(limit_frame, textvariable=self.ymin_var, width=8).grid(row=1, column=1)

        tk.Label(limit_frame, text="Y-max:").grid(row=1, column=2, padx=2)
        self.ymax_var = tk.StringVar()
        ttk.Entry(limit_frame, textvariable=self.ymax_var, width=8).grid(row=1, column=3)
        
        ttk.Button(limit_frame, text="Áp dụng cận", command=self.update_chart).grid(row=2, column=0, columnspan=4, pady=5)

        # 3. Tìm kiếm & Danh sách Y
        tk.Label(self.right_frame, text="3. Tìm kiếm tín hiệu Y:", font=('Arial', 9, 'bold')).pack(anchor="w", pady=(10,0))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_signals)
        ttk.Entry(self.right_frame, textvariable=self.search_var).pack(fill=tk.X, pady=5)

        self.listbox = tk.Listbox(self.right_frame, selectmode=tk.EXTENDED, width=35)
        for col in self.all_columns: self.listbox.insert(tk.END, col)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox.bind('<<ListboxSelect>>', lambda e: self.update_chart())

        self.canvas.mpl_connect("motion_notify_event", self.on_hover)
        self.init_annot()

    def init_annot(self):
        self.annot = self.ax.annotate("", xy=(0,0), xytext=(15,15),
                            textcoords="offset points",
                            bbox=dict(boxstyle="round", fc="yellow", alpha=0.8),
                            arrowprops=dict(arrowstyle="->"))
        self.annot.set_visible(False)

    def filter_signals(self, *args):
        query = self.search_var.get().lower()
        self.listbox.delete(0, tk.END)
        for col in self.all_columns:
            if query in col.lower(): self.listbox.insert(tk.END, col)

    def update_chart(self):
        self.ax.clear()
        self.lines = []
        x_col = self.x_axis_var.get()
        y_cols = [self.listbox.get(i) for i in self.listbox.curselection()]

        if x_col and y_cols:
            temp_df = self.df.copy()
            for col in [x_col] + y_cols:
                temp_df[col] = pd.to_numeric(temp_df[col], errors='coerce')

            for y_col in y_cols:
                if y_col == x_col: continue
                plot_data = temp_df[[x_col, y_col]].dropna().sort_values(by=x_col)
                if not plot_data.empty:
                    line, = self.ax.plot(plot_data[x_col], plot_data[y_col], 
                                         label=y_col, linewidth=1.2)
                    self.lines.append(line)

            # --- THAY ĐỔI Ở ĐÂY: Đặt tiêu đề cố định ---
            # Bạn có thể thay đổi nội dung trong dấu ngoặc kép
            self.ax.set_title("Trực Quan Hóa Dữ Liệu", fontsize=14, fontweight='bold')

            # Thiết lập Cận trục (Limits)
            try:
                if self.xmin_var.get() and self.xmax_var.get():
                    self.ax.set_xlim(float(self.xmin_var.get()), float(self.xmax_var.get()))
                if self.ymin_var.get() and self.ymax_var.get():
                    self.ax.set_ylim(float(self.ymin_var.get()), float(self.ymax_var.get()))
            except ValueError:
                pass 

            self.ax.xaxis.set_major_locator(MaxNLocator(integer=True))
            self.ax.set_xlabel(x_col)
            self.ax.set_ylabel("Giá trị")
            self.ax.grid(True, linestyle=':', alpha=0.6)
            if len(y_cols) <= 15: self.ax.legend(loc='upper right', fontsize='small')
            self.init_annot()
        
        self.fig.tight_layout()
        self.canvas.draw()

    def on_hover(self, event):
        if event.inaxes == self.ax:
            self.coord_label.config(text=f"x={event.xdata:.2f}  y={event.ydata:.2f}")
            if not self.tooltip_enabled.get():
                self.annot.set_visible(False)
                self.canvas.draw_idle()
                return

            for line in self.lines:
                cont, ind = line.contains(event)
                if cont:
                    x, y = line.get_data()
                    idx = ind["ind"][0]
                    self.annot.xy = (x[idx], y[idx])
                    self.annot.set_text(f"X: {x[idx]:.0f}\nY: {y[idx]:.2f}\n{line.get_label()}")
                    self.annot.set_visible(True)
                    self.canvas.draw_idle()
                    return
        else:
            self.coord_label.config(text="")
        
        if self.annot.get_visible():
            self.annot.set_visible(False)
            self.canvas.draw_idle()

if __name__ == "__main__":

    df_final = pd.read_excel("Test_data.xlsx")
      
    root = tk.Tk()
    app = SignalAnalyzerGUI(root, df_final)
    root.mainloop()

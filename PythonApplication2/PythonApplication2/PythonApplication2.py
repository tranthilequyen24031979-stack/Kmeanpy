import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from sklearn.cluster import KMeans
from sklearn.preprocessing import MinMaxScaler

# --- CẤU HÌNH GIAO DIỆN ---
BG, FG = "#FFFFFF", "#000000"
BTN_BG, BTN_FG = "#111111", "#FFFFFF"
FRAME_BG, HEADER_BG, HEADER_FG = "#F2F2F2", "#1A1A1A", "#FFFFFF"
CLUSTER_COLORS = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd"]

# --- LOGIC XẾP LOẠI CÁ NHÂN (NGUYÊN TẮC THẮT NÚT) ---
def tinh_loai_ca_nhan(gpa, drl):
    # Xếp hạng theo GPA (Yêu cầu của bạn)
    if   gpa >= 3.6: r_g = 5 # Xuất sắc
    elif gpa >= 3.2: r_g = 4 # Khá
    elif gpa >= 2.5: r_g = 3 # TB Khá
    elif gpa >= 2.0: r_g = 2 # Trung bình
    else:            r_g = 1 # Kém
    
    # Xếp hạng theo ĐRL (Yêu cầu của bạn)
    if   drl >= 90:  r_d = 5 # Xuất sắc
    elif drl >= 80:  r_d = 4 # Giỏi
    elif drl >= 65:  r_d = 3 # Khá
    elif drl >= 50:  r_d = 2 # Trung bình
    else:            r_d = 1 # Kém
    
    # Lấy mức thấp nhất quyết định xếp loại
    final = min(r_g, r_d)
    mapping = {5: "Xuất sắc", 4: "Giỏi/Khá", 3: "Khá/TB Khá", 2: "Trung bình", 1: "Kém"}
    return mapping[final]

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Phân Cụm & Chuẩn Hóa Dữ Liệu")
        self.state("zoomed")
        self.configure(bg=BG)
        self.df_raw = None
        self.k_var = tk.IntVar(value=3)
        self._build_ui()

    def _build_ui(self):
        tk.Label(self, text="HỆ THỐNG PHÂN CỤM VÀ XẾP LOẠI SINH VIÊN", font=("Segoe UI", 14, "bold"),
                 bg=HEADER_BG, fg=HEADER_FG, pady=10).pack(fill="x")
        
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Cột trái: Điều khiển & Bảng dữ liệu
        lf = tk.Frame(body, bg=BG, width=700)
        lf.pack(side="left", fill="both", expand=False)
        lf.pack_propagate(False)

        ctrl = tk.LabelFrame(lf, text=" Công cụ ", font=("Segoe UI", 10, "bold"), bg=FRAME_BG)
        ctrl.pack(fill="x", pady=5)
        tk.Button(ctrl, text="1. Nhập Excel", command=self._open_file, bg=BTN_BG, fg=BTN_FG).grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        tk.Label(ctrl, text="Số cụm K:", bg=FRAME_BG).grid(row=0, column=1, padx=5)
        tk.Spinbox(ctrl, from_=2, to=5, textvariable=self.k_var, width=5).grid(row=0, column=2, padx=5)
        tk.Button(ctrl, text="2. Phân cụm & Chuẩn hóa", command=self._run, bg="#28a745", fg="white").grid(row=0, column=3, sticky="ew", padx=5, pady=5)

        # Bảng hiển thị (Thêm các cột chuẩn hóa)
        cols = ("Ma", "Ten", "GPA", "DRL", "GPA_S", "DRL_S", "Cum", "Loai")
        self.tree = ttk.Treeview(lf, columns=cols, show="headings")
        headers = [("Ma", "MSSV", 60), ("Ten", "Họ Tên", 110), ("GPA", "GPA", 40), 
                   ("DRL", "ĐRL", 40), ("GPA_S", "GPA(S)", 60), ("DRL_S", "ĐRL(S)", 60),
                   ("Cum", "Cụm", 40), ("Loai", "Xếp Loại", 90)]
        for id, txt, w in headers:
            self.tree.heading(id, text=txt); self.tree.column(id, width=w, anchor="center")
        self.tree.pack(fill="both", expand=True, pady=5)

        # Cột phải: Biểu đồ
        self.fig, self.ax = plt.subplots(figsize=(5, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=body)
        self.canvas.get_tk_widget().pack(side="right", fill="both", expand=True, padx=10)

    def _open_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            df = pd.read_excel(path)
            c_map = {c: "Ma" for c in df.columns if "ma" in c.lower() or "ms" in c.lower()}
            c_map.update({c: "Ten" for c in df.columns if "ten" in c.lower()})
            c_map.update({c: "GPA" for c in df.columns if "gpa" in c.lower()})
            c_map.update({c: "DRL" for c in df.columns if "rl" in c.lower() or "rèn luyện" in c.lower()})
            self.df_raw = df.rename(columns=c_map).dropna(subset=["GPA", "DRL"])
            self._fill_table(self.df_raw)

    def _run(self):
        if self.df_raw is None: return
        k = self.k_var.get()
        X = self.df_raw[["GPA", "DRL"]].values
        
        # 1. CHUẨN HÓA DỮ LIỆU
        scaler = MinMaxScaler()
        X_sc = scaler.fit_transform(X)
        self.df_raw["GPA_S"] = np.round(X_sc[:, 0], 3)
        self.df_raw["DRL_S"] = np.round(X_sc[:, 1], 3)
        
        # 2. PHÂN CỤM K-MEANS (Dựa trên dữ liệu đã chuẩn hóa)
        km = KMeans(n_clusters=k, random_state=42, n_init=10)
        self.df_raw["Cum"] = km.fit_predict(X_sc) + 1
        
        # 3. XẾP LOẠI (Dựa trên logic yêu cầu cho từng dòng)
        self.df_raw["Loai"] = self.df_raw.apply(
            lambda r: tinh_loai_ca_nhan(r["GPA"], r["DRL"]), axis=1
        )
        
        self._fill_table(self.df_raw)
        self._draw_chart(X_sc, self.df_raw["Cum"].values - 1, km)

    def _draw_chart(self, X_sc, labels, km):
        self.ax.clear()
        for i in range(self.k_var.get()):
            mask = labels == i
            self.ax.scatter(X_sc[mask, 0], X_sc[mask, 1], c=CLUSTER_COLORS[i], label=f"Cụm {i+1}", alpha=0.7)
        
        self.ax.scatter(km.cluster_centers_[:, 0], km.cluster_centers_[:, 1], 
                        c="black", marker="*", s=150, label="Tâm")
        
        self.ax.set_xlabel("GPA Scaled (0-1)"); self.ax.set_ylabel("ĐRL Scaled (0-1)")
        self.ax.set_title(f"Phân cụm trên dữ liệu chuẩn hóa (K={self.k_var.get()})")
        self.ax.legend(); self.ax.grid(True, alpha=0.2)
        self.canvas.draw()

    def _fill_table(self, df):
        for r in self.tree.get_children(): self.tree.delete(r)
        for _, r in df.iterrows():
            self.tree.insert("", "end", values=(
                r.get("Ma",""), r.get("Ten",""), r["GPA"], r["DRL"],
                r.get("GPA_S",""), r.get("DRL_S",""),
                r.get("Cum",""), r.get("Loai","")
            ))

if __name__ == "__main__":
    App().mainloop()
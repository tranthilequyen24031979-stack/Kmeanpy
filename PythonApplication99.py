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

BG, FG = "#FFFFFF", "#000000"
BTN_BG, BTN_FG = "#111111", "#FFFFFF"
FRAME_BG, HEADER_BG, HEADER_FG = "#F2F2F2", "#1A1A1A", "#FFFFFF"
CLUSTER_COLORS = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd"]

def tinh_loai_ca_nhan(gpa, drl):
    if   gpa >= 3.6: r_g = 5
    elif gpa >= 3.2: r_g = 4
    elif gpa >= 2.5: r_g = 3
    elif gpa >= 2.0: r_g = 2
    else:            r_g = 1

    if   drl >= 90:  r_d = 5
    elif drl >= 80:  r_d = 4
    elif drl >= 65:  r_d = 3
    elif drl >= 50:  r_d = 2
    else:            r_d = 1

    final = min(r_g, r_d)
    mapping = {5: "Xuất sắc", 4: "Giỏi", 3: "Khá", 2: "Trung bình", 1: "Kém"}
    return mapping[final]

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Phân Cụm & Chuẩn Hóa Dữ Liệu Sinh Viên")
        self.state("zoomed")
        self.configure(bg=BG)
        self.df_raw = None
        self.df_result = None  
        self.k_var = tk.IntVar(value=3)
        self._build_ui()

    def _build_ui(self):
        tk.Label(self, text="HỆ THỐNG PHÂN CỤM VÀ XẾP LOẠI SINH VIÊN",
                 font=("Segoe UI", 14, "bold"),
                 bg=HEADER_BG, fg=HEADER_FG, pady=10).pack(fill="x")

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=10, pady=5)

        lf = tk.Frame(body, bg=BG, width=720)
        lf.pack(side="left", fill="both", expand=False)
        lf.pack_propagate(False)

        ctrl = tk.LabelFrame(lf, text=" Công cụ ", font=("Segoe UI", 10, "bold"), bg=FRAME_BG)
        ctrl.pack(fill="x", pady=5)

        tk.Button(ctrl, text="1. Nhập Excel", command=self._open_file,
                  bg=BTN_BG, fg=BTN_FG, width=14).grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        tk.Label(ctrl, text="Số cụm K:", bg=FRAME_BG).grid(row=0, column=1, padx=(10, 2))
        tk.Spinbox(ctrl, from_=2, to=5, textvariable=self.k_var, width=5).grid(row=0, column=2, padx=2)

        tk.Button(ctrl, text="2. Phân cụm & Chuẩn hóa", command=self._run,
                  bg="#28a745", fg="white", width=22).grid(row=0, column=3, sticky="ew", padx=5, pady=5)
     
        tk.Button(ctrl, text="3. Xuất Excel", command=self._export,
                  bg="#007bff", fg="white", width=14).grid(row=0, column=4, sticky="ew", padx=5, pady=5)

        self.status_var = tk.StringVar(value="Chưa tải file.")
        tk.Label(ctrl, textvariable=self.status_var, bg=FRAME_BG, fg="#555555",
                 font=("Segoe UI", 9)).grid(row=1, column=0, columnspan=5, sticky="w", padx=5, pady=(0, 4))

        cols = ("Ma", "Ten", "GPA", "DRL", "GPA_S", "DRL_S", "Cum", "Loai")
        self.tree = ttk.Treeview(lf, columns=cols, show="headings")
        headers = [
            ("Ma",    "MSSV",    65),
            ("Ten",   "Họ Tên", 120),
            ("GPA",   "GPA",     48),
            ("DRL",   "ĐRL",     48),
            ("GPA_S", "GPA(S)",  65),
            ("DRL_S", "ĐRL(S)",  65),
            ("Cum",   "Cụm",     45),
            ("Loai",  "Xếp Loại",95),
        ]
        for id_, txt, w in headers:
            self.tree.heading(id_, text=txt)
            self.tree.column(id_, width=w, anchor="center")

        sb_y = ttk.Scrollbar(lf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb_y.set)
        self.tree.pack(side="left", fill="both", expand=True, pady=5)
        sb_y.pack(side="left", fill="y", pady=5)

        rf = tk.Frame(body, bg=BG)
        rf.pack(side="right", fill="both", expand=True, padx=(5, 0))

        self.fig, self.ax = plt.subplots(figsize=(5, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=rf)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        stat_frame = tk.LabelFrame(rf, text=" Thống kê theo cụm ", font=("Segoe UI", 9, "bold"), bg=FRAME_BG)
        stat_frame.pack(fill="x", pady=(5, 0))
        stat_cols = ("Cum", "SoSV", "GPA_TB", "DRL_TB", "PhanBo")
        self.stat_tree = ttk.Treeview(stat_frame, columns=stat_cols, show="headings", height=5)
        stat_headers = [("Cum","Cụm",55),("SoSV","Số SV",55),("GPA_TB","GPA TB",65),("DRL_TB","ĐRL TB",65),("PhanBo","Phân bố %",80)]
        for id_, txt, w in stat_headers:
            self.stat_tree.heading(id_, text=txt)
            self.stat_tree.column(id_, width=w, anchor="center")
        self.stat_tree.pack(fill="x", padx=5, pady=5)

    def _open_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if not path:
            return
        try:
            df = pd.read_excel(path)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không đọc được file:\n{e}")
            return

        c_map = {}
        for c in df.columns:
            cl = c.lower()
            if ("ma" in cl or "ms" in cl) and "Ma" not in c_map.values():
                c_map[c] = "Ma"
            elif "ten" in cl and "Ten" not in c_map.values():
                c_map[c] = "Ten"
            elif "gpa" in cl and "GPA" not in c_map.values():
                c_map[c] = "GPA"
            elif ("rl" in cl or "rèn luyện" in cl or "diemdrl" in cl.replace(" ","")) and "DRL" not in c_map.values():
                c_map[c] = "DRL"

        df = df.rename(columns=c_map)

        missing = [c for c in ("GPA", "DRL") if c not in df.columns]
        if missing:
            messagebox.showerror("Lỗi", f"Không tìm thấy cột: {missing}\nVui lòng kiểm tra tên cột trong file Excel.")
            return

        self.df_raw = df.dropna(subset=["GPA", "DRL"]).reset_index(drop=True)
        self.df_result = None

        self._fill_table(self.df_raw, raw=True)
        self.status_var.set(f"✔ Đã tải {len(self.df_raw)} sinh viên. Nhấn 'Phân cụm' để tiếp tục.")

    def _run(self):
        if self.df_raw is None:
            messagebox.showwarning("Chưa có dữ liệu", "Vui lòng nhập file Excel trước (bước 1).")
            return

        k = self.k_var.get()
        df = self.df_raw.copy()
        X = df[["GPA", "DRL"]].values

        # CHUẨN HÓA
        scaler = MinMaxScaler()
        X_sc = scaler.fit_transform(X)
        df["GPA_S"] = np.round(X_sc[:, 0], 3)
        df["DRL_S"] = np.round(X_sc[:, 1], 3)

        # PHÂN CỤM
        km = KMeans(n_clusters=k, random_state=42, n_init=10)
        df["Cum"] = km.fit_predict(X_sc) + 1

        # XẾP LOẠI
        df["Loai"] = df.apply(lambda r: tinh_loai_ca_nhan(r["GPA"], r["DRL"]), axis=1)

        self.df_result = df
        self._fill_table(df, raw=False)
        self._draw_chart(X_sc, df["Cum"].values - 1, km, k)
        self._fill_stats(df, k)
        self.status_var.set(f"✔ Phân cụm K={k} hoàn tất. Nhấn 'Xuất Excel' để lưu kết quả.")

    def _draw_chart(self, X_sc, labels, km, k):
        self.ax.clear()
        for i in range(k):
            mask = labels == i
            self.ax.scatter(X_sc[mask, 0], X_sc[mask, 1],
                            c=CLUSTER_COLORS[i % len(CLUSTER_COLORS)],
                            label=f"Cụm {i+1}", alpha=0.75, s=55, edgecolors="white", linewidths=0.4)
        self.ax.scatter(km.cluster_centers_[:, 0], km.cluster_centers_[:, 1],
                        c="black", marker="*", s=180, zorder=5, label="Tâm cụm")
        self.ax.set_xlabel("GPA Scaled (0–1)")
        self.ax.set_ylabel("ĐRL Scaled (0–1)")
        self.ax.set_title(f"Phân cụm K-Means (K={k})")
        self.ax.legend(fontsize=8)
        self.ax.grid(True, alpha=0.25)
        self.canvas.draw()

    def _fill_table(self, df, raw=False):
        for r in self.tree.get_children():
            self.tree.delete(r)
        for _, r in df.iterrows():
            self.tree.insert("", "end", values=(
                r.get("Ma", ""),
                r.get("Ten", ""),
                round(r["GPA"], 2),
                int(r["DRL"]),
                r.get("GPA_S", "") if not raw else "",
                r.get("DRL_S", "") if not raw else "",
                r.get("Cum", "")   if not raw else "",
                r.get("Loai", "")  if not raw else "",
            ))

    def _fill_stats(self, df, k):
        for r in self.stat_tree.get_children():
            self.stat_tree.delete(r)
        total = len(df)
        for i in range(1, k + 1):
            sub = df[df["Cum"] == i]
            n = len(sub)
            gpa_tb = round(sub["GPA"].mean(), 2) if n > 0 else 0
            drl_tb = round(sub["DRL"].mean(), 1) if n > 0 else 0
            pct = round(n / total * 100, 1)
            self.stat_tree.insert("", "end", values=(f"Cụm {i}", n, gpa_tb, drl_tb, f"{pct}%"))

    def _export(self):
        if self.df_result is None:
            messagebox.showwarning("Chưa có kết quả", "Vui lòng thực hiện phân cụm trước (bước 2).")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="KetQua_PhanCum.xlsx"
        )
        if not path:
            return
        try:
            export_cols = [c for c in ["Ma", "Ten", "GPA", "DRL", "GPA_S", "DRL_S", "Cum", "Loai"] if c in self.df_result.columns]
            self.df_result[export_cols].to_excel(path, index=False)
            messagebox.showinfo("Thành công", f"Đã lưu kết quả vào:\n{path}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất file", str(e))

if __name__ == "__main__":
    App().mainloop()

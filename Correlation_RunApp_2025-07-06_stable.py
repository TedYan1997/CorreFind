
import pandas as pd
import numpy as np
import scipy.stats as stats
import dcor
import seaborn as sns
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
from datetime import datetime

# === GUI 主視窗 ===
root = tk.Tk()
root.title("Correlation Analyzer")
root.geometry("550x300")

selected_file = ""
sheet_var = tk.StringVar()
threshold_var = tk.StringVar(value="0.9")  # 預設閾值為 0.9

def select_file():
    global selected_file
    file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file:
        selected_file = file
        try:
            sheets = pd.ExcelFile(file).sheet_names
            sheet_dropdown['values'] = sheets
            sheet_var.set(sheets[0])
            lbl_file.config(text="📄 " + file.split("/")[-1])
        except Exception as e:
            messagebox.showerror("錯誤", f"無法讀取 Excel 檔案: {e}")

def plot_heatmap(df_corr, title, filename):
    plt.figure(figsize=(10, 8))
    sns.heatmap(df_corr, annot=True, fmt=".2f", cmap="coolwarm", square=True, cbar=True)
    plt.title(title)
    plt.tight_layout()
    plt.savefig(filename)
    plt.close()


def extract_filtered_pairs(df_corr, threshold):
    pairs = []
    cols = df_corr.columns
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            val = df_corr.iloc[i, j]
            if pd.notna(val) and abs(val) >= threshold:
                pairs.append({
                    "變數1": cols[i],
                    "變數2": cols[j],
                    "Correlation": round(val, 4)
                })
    result_df = pd.DataFrame(pairs)
    print(f"🔍 篩選 |corr| >= {threshold}，共找到 {len(result_df)} 組變數配對")
    return result_df
def plot_scatter_plots(df_raw, filtered_df, folder="scatter_plots"):
    if not os.path.exists(folder):
        os.makedirs(folder)
    for _, row in filtered_df.iterrows():
        var1 = row["變數1"]
        var2 = row["變數2"]
        plt.figure(figsize=(6, 4))
        plt.scatter(df_raw[var1], df_raw[var2], alpha=0.7)
        plt.xlabel(var1)
        plt.ylabel(var2)
        plt.title(f"Scatter Plot: {var1} vs {var2}\nCorr = {row['Correlation']}")
        plt.grid(True)
        filename = f"{folder}/scatter_{var1}_{var2}.png".replace(" ", "_")
        plt.tight_layout()
        plt.savefig(filename)
        plt.close()

def run_analysis():
    if not selected_file or not sheet_var.get():
        messagebox.showwarning("提醒", "請選擇 Excel 檔案與工作表")
        return

    try:
        threshold = float(threshold_var.get())
        if not (0 < threshold <= 1):
            raise ValueError("閾值必須介於 0 與 1 之間")

        now = datetime.now().strftime("%Y-%m-%d-%H%M")
        output_folder = f"correlation_result_{now}"
        os.makedirs(output_folder, exist_ok=True)

        df = pd.read_excel(selected_file, sheet_name=sheet_var.get(), header=0)
        df = df.apply(pd.to_numeric, errors='coerce').dropna()

        cols = df.columns
        n = len(cols)

        pearson_mat = np.zeros((n, n))
        spearman_mat = np.zeros((n, n))
        dcor_mat = np.zeros((n, n))

        for i in range(n):
            for j in range(n):
                x = df.iloc[:, i].values
                y = df.iloc[:, j].values
                pearson_mat[i, j] = stats.pearsonr(x, y)[0]
                spearman_mat[i, j] = stats.spearmanr(x, y)[0]
                dcor_mat[i, j] = dcor.distance_correlation(x, y)

        pearson_df = pd.DataFrame(pearson_mat, index=cols, columns=cols)
        spearman_df = pd.DataFrame(spearman_mat, index=cols, columns=cols)
        dcor_df = pd.DataFrame(dcor_mat, index=cols, columns=cols)

        pearson_filtered = extract_filtered_pairs(pearson_df, threshold)
        spearman_filtered = extract_filtered_pairs(spearman_df, threshold)
        dcor_filtered = extract_filtered_pairs(dcor_df, threshold)

        output_excel = f"{output_folder}/correlation_matrix_filtered.xlsx"
        with pd.ExcelWriter(output_excel) as writer:
            pearson_df.to_excel(writer, sheet_name="Pearson")
            spearman_df.to_excel(writer, sheet_name="Spearman")
            dcor_df.to_excel(writer, sheet_name="DistanceCorr")
            pearson_filtered.to_excel(writer, sheet_name="Pearson_Filtered", index=False)
            spearman_filtered.to_excel(writer, sheet_name="Spearman_Filtered", index=False)
            dcor_filtered.to_excel(writer, sheet_name="Distance_Filtered", index=False)

        plot_heatmap(pearson_df, "Pearson Correlation", f"{output_folder}/pearson_heatmap.png")
        plot_heatmap(spearman_df, "Spearman Correlation", f"{output_folder}/spearman_heatmap.png")
        plot_heatmap(dcor_df, "Distance Correlation", f"{output_folder}/distance_heatmap.png")

        plot_scatter_plots(df, pearson_filtered, folder=f"{output_folder}/scatter_plots/pearson")
        plot_scatter_plots(df, spearman_filtered, folder=f"{output_folder}/scatter_plots/spearman")
        plot_scatter_plots(df, dcor_filtered, folder=f"{output_folder}/scatter_plots/distance")

        messagebox.showinfo("完成", f"分析完成！\n\n結果儲存於：\n{output_folder}")

    except Exception as e:
        messagebox.showerror("錯誤", str(e))

# === GUI 元件 ===
btn_select = tk.Button(root, text="選擇 Excel 檔案", command=select_file)
btn_select.pack(pady=10)

lbl_file = tk.Label(root, text="尚未選擇檔案", wraplength=450)
lbl_file.pack()

tk.Label(root, text="選擇工作表：").pack()
sheet_dropdown = ttk.Combobox(root, textvariable=sheet_var, state="readonly", width=30)
sheet_dropdown.pack()

tk.Label(root, text="輸入 correlation 閾值 (0~1)：").pack(pady=(10, 0))
entry_threshold = tk.Entry(root, textvariable=threshold_var, justify="center", width=10)
entry_threshold.pack()

btn_run = tk.Button(root, text="執行分析", command=run_analysis, bg="lightgreen", height=2)
btn_run.pack(pady=20)

root.mainloop()

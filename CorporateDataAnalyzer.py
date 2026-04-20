import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class CorporateDataAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Corporate Data Analyzer (Report + Chart + Export)")
        self.root.geometry("1365x900")
        self.root.configure(bg="#f4f7fb")

        self.file_path = None
        self.df = None
        self.report_df = None
        self.current_figure = None
        self.chart_canvas = None
        self.text_columns = []
        self.numeric_columns = []

        self.setup_styles()
        self.build_ui()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TLabel", background="#f4f7fb", foreground="#1f2937", font=("Segoe UI", 11))
        style.configure("Title.TLabel", background="#f4f7fb", foreground="#0f172a", font=("Segoe UI", 28, "bold"))

        style.configure(
            "Section.TLabelframe",
            background="#f4f7fb",
            bordercolor="#cbd5e1",
            relief="solid"
        )
        style.configure(
            "Section.TLabelframe.Label",
            background="#f4f7fb",
            foreground="#1e293b",
            font=("Segoe UI", 12, "bold")
        )

        style.configure(
            "Treeview",
            font=("Segoe UI", 10),
            rowheight=28,
            background="white",
            fieldbackground="white",
            foreground="#111827"
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 11, "bold"),
            background="#e5e7eb",
            foreground="#111827"
        )

        style.map("Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", "#111827")])

    def build_ui(self):
        main = tk.Frame(self.root, bg="#f4f7fb")
        main.pack(fill="both", expand=True, padx=18, pady=14)

        title = ttk.Label(main, text="Corporate Data Analyzer", style="Title.TLabel")
        title.pack(pady=(10, 18))

        top_bar = tk.Frame(main, bg="#f4f7fb")
        top_bar.pack(fill="x", pady=(0, 10))

        tk.Label(top_bar, text="Select CSV/Excel:", font=("Segoe UI", 16), bg="#f4f7fb", fg="#1f2937").pack(side="left", padx=(0, 10))

        tk.Button(
            top_bar, text="Browse", width=10, font=("Segoe UI", 12, "bold"),
            bg="#f3f4f6", fg="#111827", relief="raised", bd=2, command=self.browse_file
        ).pack(side="left", padx=6)

        tk.Button(
            top_bar, text="Read", width=10, font=("Segoe UI", 12, "bold"),
            bg="#f3f4f6", fg="#111827", relief="raised", bd=2, command=self.read_file
        ).pack(side="left", padx=6)

        self.file_label = tk.Label(
            top_bar, text="No file selected", font=("Segoe UI", 12, "bold"),
            bg="#f4f7fb", fg="#312e81"
        )
        self.file_label.pack(side="left", padx=(16, 0))

        file_info_frame = ttk.LabelFrame(main, text="File Info", style="Section.TLabelframe", padding=10)
        file_info_frame.pack(fill="x", pady=(0, 14))

        self.file_info_text = tk.Text(
            file_info_frame, height=4, wrap="word", font=("Segoe UI", 11),
            bg="white", fg="#111827", relief="solid", bd=1
        )
        self.file_info_text.pack(fill="x")
        self.file_info_text.config(state="disabled")

        report_builder = ttk.LabelFrame(main, text="Build Report (GroupBy + Aggregation)", style="Section.TLabelframe", padding=12)
        report_builder.pack(fill="x", pady=(0, 14))

        row1 = tk.Frame(report_builder, bg="#f4f7fb")
        row1.pack(fill="x", pady=(0, 12))

        tk.Label(row1, text="Group By (Text column):", font=("Segoe UI", 12), bg="#f4f7fb").pack(side="left", padx=(0, 8))
        self.group_col = ttk.Combobox(row1, width=22, state="readonly")
        self.group_col.pack(side="left", padx=(0, 20))

        tk.Label(row1, text="Aggregation:", font=("Segoe UI", 12), bg="#f4f7fb").pack(side="left", padx=(0, 8))
        self.agg_combo = ttk.Combobox(row1, width=16, state="readonly", values=["sum", "mean", "average", "max", "min", "count", "median"])
        self.agg_combo.pack(side="left", padx=(0, 20))

        tk.Label(row1, text="Value (Numeric column):", font=("Segoe UI", 12), bg="#f4f7fb").pack(side="left", padx=(0, 8))
        self.value_col = ttk.Combobox(row1, width=22, state="readonly")
        self.value_col.pack(side="left")

        row2 = tk.Frame(report_builder, bg="#f4f7fb")
        row2.pack(fill="x")

        tk.Button(
            row2, text="Preview Report", width=16, font=("Segoe UI", 12, "bold"),
            bg="#2e7d32", fg="white", activebackground="#1b5e20", activeforeground="white",
            relief="raised", bd=2, command=self.preview_report
        ).pack(side="left", padx=(0, 18))

        tk.Label(row2, text="Export as:", font=("Segoe UI", 12), bg="#f4f7fb").pack(side="left", padx=(0, 8))
        self.export_combo = ttk.Combobox(row2, width=16, state="readonly", values=["Excel (.xlsx)", "CSV (.csv)"])
        self.export_combo.pack(side="left", padx=(0, 18))
        self.export_combo.set("Excel (.xlsx)")

        tk.Button(
            row2, text="Export Report", width=14, font=("Segoe UI", 12, "bold"),
            bg="#ffffff", fg="#111827", relief="raised", bd=2, command=self.export_report
        ).pack(side="left")

        chart_builder = ttk.LabelFrame(main, text="Chart Builder", style="Section.TLabelframe", padding=12)
        chart_builder.pack(fill="x", pady=(0, 14))

        tk.Label(chart_builder, text="Chart Type:", font=("Segoe UI", 12), bg="#f4f7fb").pack(side="left", padx=(0, 8))
        self.chart_combo = ttk.Combobox(chart_builder, width=12, state="readonly", values=["Bar", "Column", "Line", "Pie"])
        self.chart_combo.pack(side="left", padx=(0, 16))
        self.chart_combo.set("Bar")

        tk.Button(
            chart_builder, text="Preview Chart", width=15, font=("Segoe UI", 12, "bold"),
            bg="#1565c0", fg="white", activebackground="#0d47a1", activeforeground="white",
            relief="raised", bd=2, command=self.preview_chart
        ).pack(side="left", padx=(0, 14))

        tk.Button(
            chart_builder, text="Export Chart (PNG)", width=18, font=("Segoe UI", 12, "bold"),
            bg="#ffffff", fg="#111827", relief="raised", bd=2, command=self.export_chart
        ).pack(side="left")

        preview_area = tk.Frame(main, bg="#f4f7fb")
        preview_area.pack(fill="both", expand=True)

        left_panel = ttk.LabelFrame(preview_area, text="Report Preview", style="Section.TLabelframe", padding=10)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 8))

        tree_container = tk.Frame(left_panel, bg="white", relief="solid", bd=1)
        tree_container.pack(fill="both", expand=True)

        self.report_tree = ttk.Treeview(tree_container, columns=("Group", "Value"), show="headings")
        self.report_tree.heading("Group", text="Group")
        self.report_tree.heading("Value", text="Value")
        self.report_tree.column("Group", width=360, anchor="w")
        self.report_tree.column("Value", width=220, anchor="center")

        y_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self.report_tree.yview)
        self.report_tree.configure(yscrollcommand=y_scroll.set)

        self.report_tree.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")

        right_panel = ttk.LabelFrame(preview_area, text="Chart Preview", style="Section.TLabelframe", padding=10)
        right_panel.pack(side="left", fill="both", expand=True)

        self.chart_frame = tk.Frame(right_panel, bg="white", relief="solid", bd=1)
        self.chart_frame.pack(fill="both", expand=True)

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select CSV or Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("All Files", "*.*")
            ]
        )
        if path:
            self.file_path = path
            self.file_label.config(text=os.path.basename(path))

    def set_file_info(self, text):
        self.file_info_text.config(state="normal")
        self.file_info_text.delete("1.0", "end")
        self.file_info_text.insert("1.0", text)
        self.file_info_text.config(state="disabled")

    def detect_columns(self):
        self.text_columns = []
        self.numeric_columns = []

        for col in self.df.columns:
            series = self.df[col]
            if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
                self.text_columns.append(col)

            if pd.api.types.is_numeric_dtype(series):
                self.numeric_columns.append(col)
            else:
                converted = pd.to_numeric(series, errors="coerce")
                if converted.notna().sum() > 0:
                    self.numeric_columns.append(col)

        self.text_columns = list(dict.fromkeys(self.text_columns))
        self.numeric_columns = list(dict.fromkeys(self.numeric_columns))

    def read_file(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select a file first.")
            return

        try:
            if self.file_path.lower().endswith(".csv"):
                self.df = pd.read_csv(self.file_path)
            else:
                self.df = pd.read_excel(self.file_path)

            if self.df.empty:
                messagebox.showwarning("Warning", "The selected file is empty.")
                return

            self.detect_columns()

            self.group_col["values"] = self.text_columns
            self.value_col["values"] = self.numeric_columns
            self.group_col.set("")
            self.value_col.set("")
            self.agg_combo.set("")

            info = (
                f"Rows: {len(self.df):,}\n"
                f"Columns: {len(self.df.columns)}\n"
                f"Column Headings: {', '.join(map(str, self.df.columns.tolist()))}"
            )
            self.set_file_info(info)
            self.clear_report()
            self.clear_chart()

            messagebox.showinfo("Success", "File read successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file.\n\n{e}")

    def clear_report(self):
        self.report_df = None
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)

    def clear_chart(self):
        for widget in self.chart_frame.winfo_children():
            widget.destroy()
        if self.current_figure:
            plt.close(self.current_figure)
        self.current_figure = None
        self.chart_canvas = None

    def preview_report(self):
        if self.df is None:
            messagebox.showerror("Error", "Please read a file first.")
            return

        group_by = self.group_col.get().strip()
        agg = self.agg_combo.get().strip().lower()
        value_col = self.value_col.get().strip()

        if not group_by or not agg or not value_col:
            messagebox.showerror("Error", "Please select Group By, Aggregation, and Value Column.")
            return

        try:
            temp = self.df.copy()
            temp[value_col] = pd.to_numeric(temp[value_col], errors="coerce")
            temp = temp.dropna(subset=[group_by])

            if agg == "average":
                agg = "mean"

            result = (
                temp.groupby(group_by, dropna=False)[value_col]
                .agg(agg)
                .reset_index()
                .sort_values(by=value_col, ascending=False)
            )

            self.report_df = result.copy()
            self.report_df.columns = ["Group", "Value"]

            self.clear_report()
            self.report_df = result.copy()
            self.report_df.columns = ["Group", "Value"]

            for _, row in self.report_df.iterrows():
                group_val = "" if pd.isna(row["Group"]) else str(row["Group"])
                value = row["Value"]
                if pd.isna(value):
                    display_val = ""
                else:
                    display_val = f"{float(value):,.2f}"
                self.report_tree.insert("", "end", values=(group_val, display_val))

            self.clear_chart()
            messagebox.showinfo("Success", "Report preview generated successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Unable to generate report.\n\n{e}")

    def preview_chart(self):
        if self.report_df is None or self.report_df.empty:
            messagebox.showerror("Error", "Please preview the report first.")
            return

        self.clear_chart()

        chart_type = self.chart_combo.get().strip().lower()
        plot_df = self.report_df.copy()

        if chart_type == "pie" and len(plot_df) > 10:
            plot_df = plot_df.head(10)

        groups = plot_df["Group"].astype(str)
        values = pd.to_numeric(plot_df["Value"], errors="coerce").fillna(0)

        fig, ax = plt.subplots(figsize=(6.4, 4.8), dpi=100)
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")

        if chart_type in ["bar", "column"]:
            ax.bar(groups, values, color="#3b82f6", edgecolor="#1d4ed8")
            ax.set_xlabel("Group")
            ax.set_ylabel("Value")
        elif chart_type == "line":
            ax.plot(groups, values, marker="o", linewidth=2.5, color="#2563eb")
            ax.set_xlabel("Group")
            ax.set_ylabel("Value")
        elif chart_type == "pie":
            ax.pie(values, labels=groups, autopct="%1.1f%%", startangle=140)
            ax.axis("equal")

        ax.set_title(f"{self.chart_combo.get()} Chart", fontsize=14, fontweight="bold")
        if chart_type != "pie":
            plt.setp(ax.get_xticklabels(), rotation=35, ha="right")

        fig.tight_layout()

        self.current_figure = fig
        self.chart_canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().pack(fill="both", expand=True)

    def export_report(self):
        if self.report_df is None or self.report_df.empty:
            messagebox.showerror("Error", "Please generate the report first.")
            return

        try:
            output_dir = os.path.dirname(self.file_path) if self.file_path else os.getcwd()
            base_name = os.path.splitext(os.path.basename(self.file_path))[0] if self.file_path else "report"

            if "xlsx" in self.export_combo.get().lower():
                out_path = os.path.join(output_dir, f"{base_name}_report.xlsx")
                self.report_df.to_excel(out_path, index=False)
            else:
                out_path = os.path.join(output_dir, f"{base_name}_report.csv")
                self.report_df.to_csv(out_path, index=False)

            messagebox.showinfo("Success", f"Report exported successfully.\n\nSaved to:\n{out_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export report.\n\n{e}")

    def export_chart(self):
        if self.current_figure is None:
            messagebox.showerror("Error", "Please preview a chart first.")
            return

        try:
            output_dir = os.path.dirname(self.file_path) if self.file_path else os.getcwd()
            base_name = os.path.splitext(os.path.basename(self.file_path))[0] if self.file_path else "chart"
            chart_name = self.chart_combo.get().strip().lower()
            out_path = os.path.join(output_dir, f"{base_name}_{chart_name}_chart.png")

            self.current_figure.savefig(out_path, dpi=300, bbox_inches="tight")
            messagebox.showinfo("Success", f"Chart exported successfully.\n\nSaved to:\n{out_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export chart.\n\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = CorporateDataAnalyzer(root)
    root.mainloop()
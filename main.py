import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from pathlib import Path

from config import TEXT_COLOR_RULES
from decolorize import ColorFilter, ExcelDecolorizer, TextColorRule


class App(tk.Tk):
    VERSION = "1.0.0"
    AUTHOR = "🥷 с-нт. БОНДАРЕНКО С.В."
    TITLE = "'ІП, П. зв'язку' - Знебарвлення зведеного розкладу"
    WINDOWS_NAME = "ГЦПОС ЦК ІПтаЗ"

    def __init__(self):
        super().__init__()
        self.title(self.WINDOWS_NAME)
        self.resizable(False, False)
        self._center_window(640, 480)
        self._build_ui()

    def _center_window(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # ── Заголовок ──────────────────────────────────────────
        header = tk.Frame(self, bg="#FFFF00", height=60)
        header.pack(fill="x")
        tk.Label(
            header,
            text=self.TITLE,
            bg="#FFFF00",
            fg="black",
            font=("Segoe UI", 13, "bold"),
        ).pack(pady=15)

        # ── Вибір файлу ────────────────────────────────────────
        file_frame = tk.Frame(self, pady=16, padx=20)
        file_frame.pack(fill="x")

        tk.Label(
            file_frame, text="Файл розкладу:", font=("Segoe UI", 10)
        ).pack(anchor="w")

        row = tk.Frame(file_frame)
        row.pack(fill="x", pady=4)

        self.file_var = tk.StringVar(value="Файл не обрано")
        tk.Label(
            row,
            textvariable=self.file_var,
            font=("Segoe UI", 9),
            fg="#555",
            width=36,
            anchor="w",
            wraplength=280,
        ).pack(side="left")

        tk.Button(
            row,
            text="Огляд...",
            command=self._choose_file,
            bg="#1E7B4B",
            fg="white",
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            font=("Segoe UI", 9),
        ).pack(side="right")

        # ── Кнопка запуску ─────────────────────────────────────
        tk.Button(
            self,
            text="▶  Запустити обробку",
            command=self._run,
            bg="#1E7B4B",
            fg="white",
            relief="flat",
            padx=16,
            pady=8,
            cursor="hand2",
            font=("Segoe UI", 11, "bold"),
            activebackground="#166638",
            activeforeground="white",
        ).pack(pady=8)

        # ── Статус ─────────────────────────────────────────────
        self.status_var = tk.StringVar(value="Готово до роботи")
        tk.Label(
            self, textvariable=self.status_var, font=("Segoe UI", 9), fg="#888"
        ).pack()

        # ── Футер з авторством ─────────────────────────────────
        footer = tk.Frame(self, bg="#f0f0f0", height=30)
        footer.pack(fill="x", side="bottom")
        tk.Label(
            footer,
            text=f"v{self.VERSION}  ·  © {datetime.now().year} {self.AUTHOR}",
            bg="#f0f0f0",
            fg="#333",
            font=("Segoe UI", 8),
        ).pack(pady=6)

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="Оберіть Excel файл з розкладом",
            filetypes=[("Excel файли", "*.xlsx *.xls")],
        )
        if path:
            self.input_path = path
            self.file_var.set(Path(path).name)
            self.status_var.set("Файл обрано — натисни «Запустити»")

    def _run(self):
        if not hasattr(self, "input_path"):
            messagebox.showwarning("Увага", "Спочатку оберіть файл!")
            return

        output_path = str(
            Path(self.input_path).with_stem(
                Path(self.input_path).stem + "_decolorized"
            )
        )

        self.status_var.set("Обробка...")
        self.update()

        try:
            color_filter = ColorFilter()
            text_rules = [
                TextColorRule(target_hex=hex_color, keywords=keywords)
                for hex_color, keywords in TEXT_COLOR_RULES.items()
            ]
            decolorizer = ExcelDecolorizer(
                input_path=self.input_path,
                output_path=output_path,
                color_filter=color_filter,
                text_rules=text_rules,
            )
            decolorizer.process()

            self.status_var.set(f"Готово! Збережено: {Path(output_path).name}")
            messagebox.showinfo("Успіх", f"Файл збережено:\n{output_path}")

        except Exception as e:
            self.status_var.set("Помилка!")
            messagebox.showerror("Помилка", str(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()

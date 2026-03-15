import base64
import io
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from pathlib import Path

try:
    from PIL import Image, ImageTk

    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

from config import TEXT_COLOR_RULES, LOGO_B64
from decolorize import ColorFilter, ExcelDecolorizer, TextColorRule
from misc import format_personnel


class App(tk.Tk):
    VERSION = "1.0.0"
    AUTHOR = "🥷 с-нт. БОНДАРЕНКО С.В."
    TITLE = "Інструменти для замін в розкладі занять"
    NAME_WINDOWS = '🦔 Мультитул "Їжак"'

    def __init__(self):
        super().__init__()
        self.title(self.NAME_WINDOWS)
        self.resizable(False, False)
        self._center_window(600, 800)
        self._build_ui()

    def _center_window(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):

        # ── Логотип ────────────────────────────────────────────
        if PIL_AVAILABLE:
            bgrnd = tk.Frame(self, bg="white", height=120)
            bgrnd.pack(fill="x", side="top")
            img_data = base64.b64decode(LOGO_B64)
            pil_img = Image.open(io.BytesIO(img_data))
            pil_img.thumbnail((120, 120), Image.LANCZOS)
            self._logo = ImageTk.PhotoImage(pil_img)
            tk.Label(bgrnd, image=self._logo, bg="white").pack(pady=(0, 2))

        # ── Заголовок ──────────────────────────────────────────
        header = tk.Frame(self, bg="#1E7B4B", height=60)
        header.pack(fill="x")
        tk.Label(
            header,
            text=self.TITLE,
            bg="#1E7B4B",
            fg="white",
            font=("Segoe UI", 13, "bold"),
        ).pack(pady=15)

        # ── Вкладки ────────────────────────────────────────────
        style = ttk.Style()
        style.configure(
            "TNotebook.Tab", font=("Segoe UI", 10), padding=[12, 5]
        )

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        tab1 = tk.Frame(notebook)
        tab2 = tk.Frame(notebook)
        tab3 = tk.Frame(notebook)
        notebook.add(tab1, text="  Знебарвлення комірок  ")
        notebook.add(tab2, text="  Список ІВС  ")
        notebook.add(tab3, text="  Про додаток  ")

        self._build_tab_decolor(tab1)
        self._build_tab_personnel(tab2)
        self._build_tab_about(tab3)

        # ── Футер ──────────────────────────────────────────────
        footer = tk.Frame(self, bg="#f0f0f0")
        footer.pack(fill="x", side="bottom")
        tk.Label(
            footer,
            text=f"v{self.VERSION}  ·  ГЦПОС © {datetime.now().year} {self.AUTHOR}",
            bg="#f0f0f0",
            fg="#333",
            font=("Segoe UI", 8),
        ).pack(pady=4)

    # ── Вкладка 1: Знебарвлення ────────────────────────────────
    def _build_tab_decolor(self, parent):
        frame = tk.Frame(parent, pady=22, padx=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Файл розкладу:", font=("Segoe UI", 10)).pack(
            anchor="w"
        )

        row = tk.Frame(frame)
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
            text="📂 Вибрати файл...",
            command=self._choose_file,
            bg="#1E7B4B",
            fg="white",
            relief="flat",
            padx=15,
            pady=8,
            cursor="hand2",
            font=("Segoe UI", 9),
        ).pack(side="right")

        tk.Button(
            frame,
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
        ).pack(pady=12)

        self.status_var = tk.StringVar(value="Готово до роботи")
        tk.Label(
            frame,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            fg="#888",
        ).pack()

    # ── Вкладка 2: Список особового складу ────────────────────
    def _build_tab_personnel(self, parent):
        frame = tk.Frame(parent, pady=12, padx=16)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text="Вставте список (звання прізвище ініціали):",
            font=("Segoe UI", 10),
        ).pack(anchor="w")

        self.input_text = tk.Text(
            frame,
            height=7,
            font=("Segoe UI", 10),
            relief="solid",
            bd=1,
            wrap="word",
        )
        self.input_text.pack(fill="x", pady=(4, 8))

        tk.Button(
            frame,
            text="⟳  Перетворити",
            command=self._convert_personnel,
            bg="#1E7B4B",
            fg="white",
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            font=("Segoe UI", 10, "bold"),
            activebackground="#166638",
            activeforeground="white",
        ).pack()

        tk.Label(frame, text="Результат:", font=("Segoe UI", 10)).pack(
            anchor="w", pady=(10, 0)
        )

        self.output_text = tk.Text(
            frame,
            height=4,
            font=("Segoe UI", 10),
            relief="solid",
            bd=1,
            wrap="word",
            state="disabled",
            bg="#f7f7f7",
        )
        self.output_text.pack(fill="x", pady=(4, 8))

        tk.Button(
            frame,
            text="📋  Копіювати",
            command=self._copy_result,
            bg="#4472C4",
            fg="white",
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            font=("Segoe UI", 10),
            activebackground="#2E5BA8",
            activeforeground="white",
        ).pack()

    # ── Вкладка 3: Про додаток ────────────────────────────────
    def _build_tab_about(self, parent):
        frame = tk.Frame(parent, pady=20, padx=24)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text='Мультитул "Їжак"',
            font=("Segoe UI", 15, "bold"),
            fg="#1E7B4B",
        ).pack(anchor="w")

        tk.Label(
            frame,
            text=f"Версія {self.VERSION}",
            font=("Segoe UI", 9),
            fg="#888",
        ).pack(anchor="w", pady=(0, 12))

        desc = (
            "Програмний інструмент для автоматизації рутинних задач\n"
            "при роботі з замінами в розкладі та документами підрозділу.\n"
        )
        tk.Label(
            frame,
            text=desc,
            font=("Segoe UI", 10),
            justify="left",
            wraplength=380,
        ).pack(anchor="w")

        # Функції
        tk.Label(
            frame, text="Поточні функції:", font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", pady=(8, 2))

        functions = [
            "🔸  Знебарвлення Excel-таблиць зі збереженням вибраних кольорів, які стосуються підрозділу",
            "🔸  Автоматичне зафарбовування комірок за ключовими словами, тем занять підрозділу ",
            "🔸  Форматування списку інструкторсько-викладацького складу в один рядок",
        ]
        for f in functions:
            tk.Label(
                frame,
                text=f,
                font=("Segoe UI", 10),
                justify="left",
                wraplength=380,
            ).pack(anchor="w", pady=1)

        tk.Label(
            frame,
            text="\nДодаток розроблено для потреб ГЦПОС ЦК ІПтаЗ.\nМожливе масштабування функціоналу в наступних версіях.",
            font=("Segoe UI", 9),
            fg="#555",
            justify="left",
            wraplength=380,
        ).pack(anchor="w")

        # Розділювач
        ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=14)

        # Автор
        tk.Label(
            frame,
            text="Розробник:",
            font=("Segoe UI", 9, "bold"),
            fg="#555",
        ).pack(anchor="w")
        tk.Label(
            frame,
            text="с-нт. БОНДАРЕНКО С.В.",
            font=("Segoe UI", 10),
            fg="#1E7B4B",
        ).pack(anchor="w")

    def _convert_personnel(self):
        raw = self.input_text.get("1.0", "end")
        result = format_personnel(raw)
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", "end")
        self.output_text.insert("1.0", result)
        self.output_text.config(state="disabled")

    def _copy_result(self):
        result = self.output_text.get("1.0", "end").strip()
        if result:
            self.clipboard_clear()
            self.clipboard_append(result)
            messagebox.showinfo(
                "Скопійовано", "Текст скопійовано в буфер обміну!"
            )
        else:
            messagebox.showwarning("Увага", "Немає тексту для копіювання.")

    # ── Знебарвлення ───────────────────────────────────────────
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
            Path("output_files").mkdir(exist_ok=True)
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

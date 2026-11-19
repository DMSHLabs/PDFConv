"""Окно-конвертер: перетащите файлы .md или .tex — рядом появятся PDF.

Запуск без консоли:  pyw -3 pdfconv.pyw
Запуск с файлами:    pyw -3 pdfconv.pyw отчёт.md статья.tex
Проверка окружения:  py -3 convert.py --check
"""
from __future__ import annotations

import json
import os
import queue
import re
import sys
import threading
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path

# Данные проекта (bin, lib, res, pdfconv.json) лежат в корне, а сам код — в src, поэтому
# различаем каталог кода и корень: в собранном exe корень — папка с exe.
CODE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else CODE_DIR.parent
# Локальный lib нужен только при запуске через Python: в собранном exe зависимости уже внутри,
# а лишний путь в sys.path только мешает импортам.
if not getattr(sys, "frozen", False):
    sys.path.insert(0, str(PROJECT_DIR / "lib"))

import tkinter as tk  # noqa: E402
from tkinter import filedialog, ttk  # noqa: E402

import convert  # noqa: E402

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:  # noqa: BLE001 — без библиотеки окно работает, но без перетаскивания
    DND_AVAILABLE = False

FILE_TYPES = [("Markdown, LaTeX и Word", "*.md *.markdown *.tex *.latex *.docx"),
              ("Все файлы", "*.*")]
CONFIG_PATH = PROJECT_DIR / "pdfconv.json"  # локальные настройки окна, в git не хранятся
ICON_PATH = PROJECT_DIR / "res" / "PDFConv.ico"
DEFAULT_SIZE = (980, 760)
BG = "#f6f7f9"
OK_COLOR = "#137333"
WARN_COLOR = "#b06000"
ERR_COLOR = "#c5221f"
LINK_COLOR = "#0b57d0"


def load_config() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_config(config: dict) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def default_text_dir() -> Path:
    """Куда сохранять текст, вставленный из буфера: Документы\\PDFConv."""
    documents = Path.home() / "Documents"
    return (documents if documents.is_dir() else Path.home()) / "PDFConv"


def text_file_name(text: str, directory: Path) -> Path:
    """Имя файла из первого заголовка, иначе «Текст»; существующее имя не перетираем."""
    heading = next((line.lstrip("#").strip() for line in text.splitlines()
                    if line.strip().startswith("#")), "")
    base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", heading).strip(" .")[:60] or "Текст"
    candidate = directory / (base + ".md")
    if candidate.exists():
        candidate = directory / ("%s-%s.md" % (base, datetime.now().strftime("%Y%m%d-%H%M%S")))
    return candidate


class DropZone(tk.Canvas):
    """Пунктирная область, которая принимает перетаскиваемые файлы."""

    def __init__(self, master, on_files, dnd_available: bool):
        super().__init__(master, height=150, bg=BG, highlightthickness=0)
        self.on_files = on_files
        self.dnd_available = dnd_available
        self.text = ("Перетащите сюда файлы .md, .markdown, .tex, .latex, .docx"
                     if dnd_available else
                     "Перетаскивание недоступно: не установлена tkinterdnd2.\n"
                     "Нажмите «Выбрать файлы…» — подробности в журнале ниже")
        self.hint = "Файлы можно перетаскивать и по одному, и сразу группой" if dnd_available else ""
        self.active = False
        self.bind("<Configure>", lambda _event: self.redraw())
        if dnd_available:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<DropEnter>>", self._on_enter)
            self.dnd_bind("<<DropLeave>>", self._on_leave)
            self.dnd_bind("<<Drop>>", self._on_drop)

    def redraw(self):
        self.delete("all")
        width, height = self.winfo_width(), self.winfo_height()
        color = LINK_COLOR if self.active else "#9aa0a6"
        self.create_rectangle(4, 4, width - 4, height - 4, outline=color, width=2, dash=(6, 4))
        self.create_text(width / 2, height / 2 - (12 if self.hint else 0), text=self.text,
                         font=("Segoe UI", 12), fill="#202124", justify="center", width=width - 40)
        if self.hint:
            self.create_text(width / 2, height / 2 + 24, text=self.hint,
                             font=("Segoe UI", 9), fill="#5f6368", justify="center")

    def _set_active(self, active: bool):
        self.active = active
        self.redraw()

    def _on_enter(self, _event):
        self._set_active(True)
        return "copy"

    def _on_leave(self, _event):
        self._set_active(False)
        return "copy"

    def _on_drop(self, event):
        self._set_active(False)
        paths = [Path(item) for item in self.tk.splitlist(event.data)]
        self.on_files(paths)
        return "copy"


class ConverterApp:
    def __init__(self, root: tk.Tk, initial: list[Path] | None = None):
        self.root = root
        self.queue: queue.Queue = queue.Queue()
        self.running = False
        self.env_running = False
        self.last_output_dir: Path | None = None
        self.results: list[convert.Result] = []

        root.title("Конвертер Markdown / LaTeX в PDF и DOCX")
        self._fit_to_screen()
        self._set_icon()
        root.protocol("WM_DELETE_WINDOW", self._on_close)
        root.configure(bg=BG)

        self._build_header()
        self._build_panes()
        self._build_workspace()
        self._build_actions()
        self._build_log()

        self.refresh_environment_soon()
        self.root.after(150, self._restore_sash)
        self._queue_poll = self.root.after(80, self._drain_queue)
        if initial:
            self.root.after(300, lambda: self.start_batch(initial))

    def _restore_sash(self):
        """Положение сплиттера: запомненное или по умолчанию (вкладкам — 240 px)."""
        stored = load_config().get("sash_pos")
        available = self.paned.winfo_height()
        position = int(stored) if isinstance(stored, int) else 240
        position = max(150, min(position, max(150, available - 240)))
        try:
            self.paned.sashpos(0, position)
        except tk.TclError:  # окно закрыли раньше, чем оно успело отрисоваться
            pass

    def refresh_environment_soon(self):
        """Проверки окружения запускают kpsewhich, поэтому рисуем окно раньше них."""
        self.root.after(60, self.refresh_environment)

    def _fit_to_screen(self):
        """Размер окна: запомненный в прошлый раз или по умолчанию, но не больше экрана."""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width, height = DEFAULT_SIZE
        stored = str(load_config().get("window_size", ""))
        match = re.fullmatch(r"(\d{3,5})x(\d{3,5})", stored)
        if match:  # пользователь растянул окно — уважаем это
            width, height = int(match.group(1)), int(match.group(2))
        width = max(780, min(width, screen_width - 60))
        height = max(560, min(height, screen_height - 90))
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 3)
        self.root.geometry("%dx%d+%d+%d" % (width, height, x, y))
        self.root.minsize(min(820, width), min(560, height))

    def _on_close(self):
        if self._queue_poll:  # иначе отложенный опрос очереди ругается на закрытое окно
            self.root.after_cancel(self._queue_poll)
            self._queue_poll = None
        config = load_config()
        config["window_size"] = "%dx%d" % (self.root.winfo_width(), self.root.winfo_height())
        try:
            config["sash_pos"] = int(self.paned.sashpos(0))
        except (tk.TclError, AttributeError):
            pass
        save_config(config)
        self.root.destroy()

    def _set_icon(self):
        """Иконка окна — та же, что у собранного exe."""
        if not ICON_PATH.exists():
            return
        try:
            self.root.iconbitmap(default=str(ICON_PATH))
        except tk.TclError:
            pass

    # --- построение окна -----------------------------------------------------

    def _build_header(self):
        frame = tk.Frame(self.root, bg=BG)
        frame.pack(fill="x", padx=14, pady=(12, 6))
        tk.Label(frame, text="Конвертер Markdown и LaTeX в PDF и DOCX", bg=BG,
                 font=("Segoe UI", 15, "bold")).pack(anchor="w")
        self.subtitle = tk.Label(
            frame, text="Сборка через локальные pandoc и MiKTeX: формулы, таблицы, "
                        "кириллица и китайский текст. Перетащите файлы или вставьте "
                        "markdown текстом — результат появится рядом с исходником "
                        "или в папке для текста.",
            bg=BG, fg="#5f6368", font=("Segoe UI", 9), justify="left", anchor="w")
        self.subtitle.pack(fill="x")
        # подзаголовок переносится по ширине окна, а не обрезается по краям
        frame.bind("<Configure>",
                   lambda event: self.subtitle.configure(wraplength=max(240, event.width - 8)))

    def _build_panes(self):
        """Верх (вкладки) и низ (кнопки + журнал) разделены сплиттером."""
        self.paned = ttk.PanedWindow(self.root, orient="vertical")
        self.paned.pack(fill="both", expand=True, padx=14, pady=(0, 12))
        self.top_pane = tk.Frame(self.paned, bg=BG)
        self.bottom_pane = tk.Frame(self.paned, bg=BG)
        self.paned.add(self.top_pane, weight=0)  # лишнюю высоту забирает нижняя панель
        self.paned.add(self.bottom_pane, weight=1)

    def _build_workspace(self):
        """Две вкладки: файлы перетаскиваем, текст вставляем из буфера."""
        self.notebook = ttk.Notebook(self.top_pane)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.add(self._build_files_tab(), text="  Файлы  ")
        self.notebook.add(self._build_text_tab(), text="  Текст  ")

    def _build_files_tab(self):
        tab = tk.Frame(self.notebook, bg=BG)
        self.drop_zone = DropZone(tab, self.start_batch, self.dnd_available)
        self.drop_zone.pack(fill="both", expand=True, padx=2, pady=(4, 2))
        tk.Label(tab, text="Либо вставьте markdown из буфера обмена на вкладке «Текст».",
                 bg=BG, fg="#5f6368", font=("Segoe UI", 9)).pack(anchor="w", padx=4, pady=(0, 4))
        if self.dnd_available:
            # файлы принимает вся вкладка, а не только пунктирный прямоугольник
            tab.drop_target_register(DND_FILES)
            tab.dnd_bind("<<Drop>>", self._on_tab_drop)
        return tab

    def _on_tab_drop(self, event):
        """Файлы, брошенные на вкладку «Файлы» мимо пунктирной зоны."""
        paths = [Path(item) for item in self.tk.splitlist(event.data)]
        self.start_batch(paths)
        return "copy"

    def _build_text_tab(self):
        tab = tk.Frame(self.notebook, bg=BG)
        toolbar = tk.Frame(tab, bg=BG)
        toolbar.pack(fill="x", padx=2, pady=(4, 2))
        ttk.Button(toolbar, text="Вставить из буфера", command=self.paste_from_clipboard
                   ).pack(side="left")
        ttk.Button(toolbar, text="Очистить", command=self.clear_text).pack(side="left", padx=6)
        self.text_button = ttk.Button(toolbar, text="Конвертировать текст",
                                     command=self.convert_text)
        self.text_button.pack(side="left")

        self.text_button_out = ttk.Button(toolbar, text="Изменить папку…",
                                         command=self.choose_text_dir)
        self.text_button_out.pack(side="right")
        self.text_dir_label = tk.Label(toolbar, text="", bg=BG, fg="#5f6368",
                                       font=("Segoe UI", 9))
        self.text_dir_label.pack(side="right", padx=8)
        tk.Label(toolbar, text="Сохранять в:", bg=BG, fg="#202124",
                 font=("Segoe UI", 9, "bold")).pack(side="right")

        text_frame = tk.Frame(tab, bg=BG)
        text_frame.pack(fill="both", expand=True, padx=2, pady=(0, 4))
        self.text_area = tk.Text(text_frame, wrap="word", height=9, font=("Consolas", 10),
                                 bg="white", relief="solid", bd=1, padx=8, pady=6, undo=True)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.text_area.yview)
        self.text_area.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.text_area.pack(side="left", fill="both", expand=True)
        # правый клик — «Вставить» в позицию курсора (кнопка «Вставить из буфера»
        # заменяет всё поле целиком, а контекстное меню вставляет как редактор)
        self.text_menu = tk.Menu(self.text_area, tearoff=0)
        self.text_menu.add_command(label="Вставить", command=self.paste_at_cursor)
        self.text_area.bind("<Button-3>", self._show_text_menu)
        hint = tk.Label(tab, text="Работает Ctrl+V и правка прямо в поле, правый клик — "
                                "«Вставить» в позицию курсора. Формулы сохраняются "
                                "как есть: $…$ в строке, $$…$$ отдельным блоком. Имя файла "
                                "берётся из первого заголовка #.",
                       bg=BG, fg="#5f6368", font=("Segoe UI", 9), justify="left", anchor="w")
        hint.pack(fill="x", padx=4, pady=(0, 4))
        tab.bind("<Configure>",
                 lambda event: hint.configure(wraplength=max(240, event.width - 12)))
        self.update_text_dir_label()
        return tab

    def _build_actions(self):
        frame = tk.Frame(self.bottom_pane, bg=BG)
        frame.pack(fill="x", pady=(8, 0))
        self.pick_button = ttk.Button(frame, text="Выбрать файлы…", command=self.pick_files)
        self.pick_button.pack(side="left")
        self.folder_button = ttk.Button(frame, text="Открыть папку с результатами",
                                       command=self.open_output_folder, state="disabled")
        self.folder_button.pack(side="left", padx=6)
        self.recheck_button = ttk.Button(frame, text="Проверить окружение снова",
                                        command=self.refresh_environment)
        self.recheck_button.pack(side="left")
        ttk.Button(frame, text="Очистить журнал", command=self.clear_log).pack(side="right")

        formats_frame = tk.Frame(self.bottom_pane, bg=BG)
        formats_frame.pack(fill="x", pady=(8, 0))
        tk.Label(formats_frame, text="Convert to:", bg=BG, fg="#202124",
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        self.format_vars = {}
        defaults = {"pdf": True, "docx": True, "md": False}  # md — под .docx-вход, включается сам
        for name, label in (("pdf", "pdf"), ("docx", "docx"), ("md", "md")):
            variable = tk.BooleanVar(value=defaults[name])
            self.format_vars[name] = variable
            ttk.Checkbutton(formats_frame, text=label, variable=variable).pack(side="left", padx=(6, 0))
        tk.Label(formats_frame, text="(каждый отмеченный формат создаётся рядом с исходником)",
                 bg=BG, fg="#5f6368", font=("Segoe UI", 9)).pack(side="left", padx=10)

        progress_frame = tk.Frame(self.bottom_pane, bg=BG)
        progress_frame.pack(fill="x", pady=(8, 4))
        self.progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress.pack(fill="x")
        self.progress_label = tk.Label(progress_frame, text="Готов к работе", bg=BG,
                                       fg="#5f6368", font=("Segoe UI", 9), anchor="w")
        self.progress_label.pack(fill="x", pady=(2, 0))

    def _build_log(self):
        frame = tk.LabelFrame(self.bottom_pane, text=" Журнал ", bg=BG,
                              font=("Segoe UI", 10, "bold"), fg="#202124",
                              labelanchor="nw", bd=1, relief="groove")
        frame.pack(fill="both", expand=True, pady=(4, 0))
        inner = tk.Frame(frame, bg=BG)
        inner.pack(fill="both", expand=True, padx=10, pady=(4, 8))
        self.log = tk.Text(inner, wrap="word", height=14, font=("Consolas", 10),
                           bg="white", relief="solid", bd=1, padx=8, pady=6)
        scrollbar = ttk.Scrollbar(inner, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        self.log.tag_configure("ok", foreground=OK_COLOR)
        self.log.tag_configure("warn", foreground=WARN_COLOR)
        self.log.tag_configure("err", foreground=ERR_COLOR)
        self.log.tag_configure("note", foreground="#5f6368")
        self.log.tag_configure("head", font=("Consolas", 10, "bold"))
        self.log.tag_configure("summary", font=("Consolas", 10, "bold"), foreground="#202124")
        self.log.tag_configure("link", foreground=LINK_COLOR, underline=True)
        self.log.tag_bind("link", "<Button-1>", self._open_link)
        self.log.tag_bind("link", "<Enter>", lambda _event: self.log.configure(cursor="hand2"))
        self.log.tag_bind("link", "<Leave>", lambda _event: self.log.configure(cursor=""))
        self.log.configure(state="disabled")
        # правый клик в журнале: скопировать выделенное или выделить всё
        self.log_menu = tk.Menu(self.log, tearoff=0)
        self.log_menu.add_command(label="Копировать", command=self.copy_log_selection)
        self.log_menu.add_command(label="Выделить всё", command=self.select_all_log)
        self.log.bind("<Button-3>", self._show_log_menu)

    def _open_link(self, event):
        """Ссылка в журнале («откуда взять») открывается по клику."""
        index = self.log.index("@%d,%d" % (event.x, event.y))
        ranges = self.log.tag_ranges("link")
        for start, end in zip(ranges[::2], ranges[1::2]):
            if self.log.compare(start, "<=", index) and self.log.compare(index, "<", end):
                webbrowser.open(self.log.get(start, end).strip())
                return

    @property
    def dnd_available(self) -> bool:
        return DND_AVAILABLE

    # --- окружение ----------------------------------------------------------

    def refresh_environment(self):
        """Проверки окружения идут в журнал: при запуске и по кнопке «Проверить снова».

        Пробная сборка LaTeX занимает секунды, поэтому проверяем в отдельном потоке —
        иначе окно замирает на это время.
        """
        if self.env_running:
            return
        self.env_running = True
        self.recheck_button.configure(state="disabled")
        self.write_log("Проверяю окружение…", "note")
        threading.Thread(target=self._environment_worker, daemon=True).start()

    def _environment_worker(self):
        try:
            self.queue.put(("env", 0, convert.check_environment()))
        except Exception as error:  # noqa: BLE001
            self.queue.put(("env_failed", 0, error))

    def _render_environment(self, checks):
        self.env_running = False
        self.recheck_button.configure(state="normal")
        missing_required = [check for check in checks if not check.ok and check.required]
        self.write_log("Проверка окружения:", "head")
        for check in checks:
            self._log_dependency(check)
        if missing_required:
            self.write_log("Нужно доустановить: %s — без этого сборка не заработает."
                           % ", ".join(check.title.split(" (")[0] for check in missing_required),
                           "err")
        else:
            self.write_log("Окружение в порядке — можно конвертировать.", "ok")
        self.write_log("", "")

    def _log_dependency(self, check: convert.Dependency):
        mark = "✓" if check.ok else ("✗" if check.required else "!")
        tag = "ok" if check.ok else ("err" if check.required else "warn")
        self.write_log("  %s %s: %s" % (mark, check.title, check.detail), tag)
        if not check.ok and check.fix:
            self.write_log("      что сделать: %s" % check.fix, "note")
        if not check.ok and check.url:
            self.write_log("      откуда взять: %s" % check.url, "link")

    # --- действия -----------------------------------------------------------

    def pick_files(self):
        paths = filedialog.askopenfilenames(title="Выберите файлы для конвертации",
                                            filetypes=FILE_TYPES)
        if paths:
            self.start_batch([Path(item) for item in paths])

    # --- вкладка «Текст» ----------------------------------------------------

    def text_output_dir(self) -> Path:
        """Папка для результатов, полученных из вставленного текста (запоминается)."""
        configured = load_config().get("text_output_dir")
        return Path(configured) if configured else default_text_dir()

    def update_text_dir_label(self):
        self.text_dir_label.configure(text=str(self.text_output_dir()))

    def choose_text_dir(self):
        directory = filedialog.askdirectory(title="Куда сохранять результаты из текста",
                                            initialdir=str(self.text_output_dir()))
        if not directory:
            return
        config = load_config()
        config["text_output_dir"] = directory
        save_config(config)
        self.update_text_dir_label()
        self.write_log("Результаты из текста будут сохраняться в %s" % directory, "note")

    def paste_from_clipboard(self):
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            self.write_log("В буфере обмена нет текста (например, там картинка) — "
                           "скопируйте markdown и попробуйте снова.", "err")
            return
        if not text.strip():
            self.write_log("Буфер обмена пуст.", "note")
            return
        self.text_area.delete("1.0", "end")
        self.text_area.insert("1.0", text)
        self.write_log("Вставлено из буфера: %d символов. Нажмите «Конвертировать текст», "
                       "если нужно что-то поправить — правьте прямо в поле." % len(text), "ok")

    def paste_at_cursor(self):
        """Вставка из буфера в позицию курсора — правый клик в поле текста."""
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            self.write_log("В буфере обмена нет текста (например, там картинка) — "
                           "скопируйте markdown и попробуйте снова.", "err")
            return
        self.text_area.insert(tk.INSERT, text)

    def _show_text_menu(self, event):
        self.text_area.focus_set()
        try:
            self.text_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.text_menu.grab_release()

    def copy_log_selection(self):
        """Скопировать выделенное в журнале — правый клик → «Копировать»."""
        try:
            text = self.log.get(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            return  # ничего не выделено
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    def select_all_log(self):
        self.log.tag_add(tk.SEL, "1.0", "end")

    def _show_log_menu(self, event):
        try:
            self.log_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.log_menu.grab_release()

    def clear_text(self):
        self.text_area.delete("1.0", "end")

    def convert_text(self):
        if self.running:
            self.write_log("Конвертация уже идёт — дождитесь окончания.", "note")
            return
        text = self.text_area.get("1.0", "end").strip()
        if not text:
            self.write_log("Текст пуст: вставьте markdown в поле или перетащите файл "
                           "на вкладке «Файлы».", "err")
            return
        directory = self.text_output_dir()
        try:
            directory.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            self.write_log("Не удалось создать папку %s: %s. Выберите другую кнопкой "
                           "«Изменить папку…»." % (directory, error), "err")
            return
        path = text_file_name(text, directory)
        try:
            path.write_text(text if text.endswith("\n") else text + "\n", encoding="utf-8")
        except OSError as error:
            self.write_log("Не удалось записать %s: %s" % (path, error), "err")
            return
        self.write_log("Текст сохранён как %s" % path, "note")
        self.start_batch([path])

    def open_output_folder(self):
        if self.last_output_dir:
            webbrowser.open(self.last_output_dir.as_uri())

    def clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def write_log(self, message: str, tag: str = ""):
        if os.environ.get("PDFCONV_DEBUG"):
            # отладка собранного exe (у него нет консоли): журнал дублируется в файл
            try:
                with open(Path.home() / "pdfconv-debug.log", "a", encoding="utf-8") as debug:
                    debug.write(datetime.now().strftime("%H:%M:%S ") + message + "\n")
            except OSError:
                pass
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def selected_formats(self) -> tuple[str, ...]:
        return tuple(name for name, variable in self.format_vars.items() if variable.get())

    def start_batch(self, paths: list[Path]):
        if self.running:
            self.write_log("Конвертация уже идёт — дождитесь окончания.", "note")
            return
        files = [Path(path) for path in paths]
        if not files:
            return
        formats = self.selected_formats()
        if any(path.suffix.lower() == ".docx" for path in files) and "md" not in formats:
            # смысл .docx-входа — восстановленный markdown с формулами LaTeX
            formats = formats + ("md",)
            self.write_log("Для .docx добавлен формат md: восстановленный текст с формулами "
                           "LaTeX.", "note")
        if not formats:
            self.write_log("Отметьте хотя бы один формат в строке «Convert to:» — "
                           "pdf, docx, md или их сочетание.", "err")
            return
        self.running = True
        self.results = []
        self.pick_button.configure(state="disabled")
        self.progress.configure(maximum=len(files), value=0)
        self.write_log("Начинаю: %d файл(ов), форматы: %s" % (len(files), ", ".join(formats)), "head")
        threading.Thread(target=self._worker, args=(files, formats), daemon=True).start()

    def _worker(self, files: list[Path], formats: tuple[str, ...]):
        for index, path in enumerate(files, start=1):
            self.queue.put(("start", index, path))
            try:
                result = convert.convert_file(path, formats=formats)
            except Exception:  # noqa: BLE001
                result = convert.Result(source=path, requested=formats)
                result.errors.append("сбой: " + traceback.format_exc(limit=1).strip())
            self.queue.put(("done", index, result))
        self.queue.put(("finished", len(files), None))

    def _drain_queue(self):
        try:
            while True:
                kind, index, payload = self.queue.get_nowait()
                if kind == "start":
                    self.progress_label.configure(
                        text="Конвертируется (%d): %s" % (index, Path(payload).name))
                elif kind == "done":
                    self.results.append(payload)
                    self._show_result(payload)
                    self.progress.configure(value=index)
                elif kind == "env":
                    self._render_environment(payload)
                elif kind == "env_failed":
                    self.env_running = False
                    self.recheck_button.configure(state="normal")
                    self.write_log("Не удалось выполнить проверку окружения: %s" % payload, "err")
                elif kind == "finished":
                    self._finish_batch()
        except queue.Empty:
            pass
        self._queue_poll = self.root.after(80, self._drain_queue)

    def _show_result(self, result: convert.Result):
        name = result.source.name
        if result.status == "error":
            self.write_log("✗ %s — ошибка" % name, "err")
        elif result.status == "warnings":
            self.write_log("! %s — готово, есть предупреждения" % name, "warn")
        else:
            pages = " (%d стр.)" % result.pages if result.pages else ""
            self.write_log("✓ %s — готово%s" % (name, pages), "ok")
        for format_name in result.requested:
            produced = getattr(result, format_name)
            if produced:
                self.write_log("    → %s" % produced, "note")
        for message in result.errors:
            self.write_log("    ошибка: %s" % message, "err")
        for message in result.warnings:
            self.write_log("    предупреждение: %s" % message, "warn")
        for message in result.notes:
            self.write_log("    примечание: %s" % message, "note")
        for format_name in result.requested:
            produced = getattr(result, format_name)
            if produced:
                self.last_output_dir = produced.parent
                break

    def _finish_batch(self):
        self.running = False
        self.pick_button.configure(state="normal")
        self.progress_label.configure(text="Готово: %d файл(ов)" % len(self.results))
        if self.last_output_dir:
            self.folder_button.configure(state="normal")
        self.write_log("", "")
        self.write_log("Итог:", "head")
        for line in convert.summarize(self.results):
            self.write_log("  " + line, "summary")
        failed = [result for result in self.results if result.status == "error"]
        if failed:
            self.write_log("", "")
            self.write_log("Что не получилось:", "head")
            for result in failed:
                reason = result.errors[0] if result.errors else "PDF не создан"
                self.write_log("  %s — %s" % (result.source.name, reason), "err")


def _wide_argv() -> list[str]:
    """Аргументы командной строки как Unicode.

    Собранный noconsole-exe может получить argv в ANSI-кодировке: кириллица в именах
    файлов превращается в мусор ещё до python (файл «не найден» при живом файле).
    Читаем командную строку как широкий Unicode напрямую; при неудаче — sys.argv.
    """
    import ctypes
    try:
        number = ctypes.c_int(0)
        get_argv = ctypes.windll.shell32.CommandLineToArgvW
        get_argv.argtypes = [ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_int)]
        get_argv.restype = ctypes.POINTER(ctypes.c_wchar_p)
        pointer = get_argv(ctypes.windll.kernel32.GetCommandLineW(), ctypes.byref(number))
        if not pointer:
            return sys.argv
        try:
            arguments = [pointer[index] for index in range(number.value)]
        finally:
            ctypes.windll.kernel32.LocalFree(pointer)
        return arguments
    except Exception:  # noqa: BLE001
        return sys.argv


def main() -> int:
    files = [Path(argument) for argument in _wide_argv()[1:]]
    if os.environ.get("PDFCONV_DEBUG"):
        try:
            with open(Path.home() / "pdfconv-debug.log", "a", encoding="utf-8") as debug:
                debug.write("argv: %r\n" % (sys.argv,))
                debug.write("cwd: %r\n" % os.getcwd())
                for item in files:
                    debug.write("файл %r exists=%s\n" % (str(item), item.exists()))
        except OSError:
            pass
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except tk.TclError:
        pass
    ConverterApp(root, initial=files)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
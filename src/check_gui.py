"""Проверка окна без показа: галочки форматов, защита от пустого выбора, drag&drop.

Запуск: C:\\Tools\\Scripts\\Python314\\python.exe check_gui.py
Окно создаётся и сразу уничтожается, mainloop не запускается. Возвращает 1, если
какая-то проверка не прошла — годится как быстрый тест после правок в pdfconv.pyw.
"""
import importlib.util
import sys
import time
from pathlib import Path

# Код лежит в src, зависимости и результаты — в корне проекта.
CODE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = CODE_DIR.parent
sys.path.insert(0, str(CODE_DIR))
sys.path.insert(0, str(PROJECT_DIR / "lib"))

import convert  # noqa: E402

convert.utf8_console()

spec = importlib.util.spec_from_file_location("pdfconv", CODE_DIR / "pdfconv.pyw")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

failures: list[str] = []


def check(title: str, condition: bool) -> None:
    print("%s %s" % ("OK  " if condition else "ПРОВАЛ", title))
    if not condition:
        failures.append(title)


root = module.TkinterDnD.Tk() if module.DND_AVAILABLE else module.tk.Tk()
app = module.ConverterApp(root)
root.update()

check("drag&drop подключён (tkinterdnd2)", getattr(app.drop_zone, "dnd_available", False))
files_tab = app.notebook.nametowidget(app.notebook.tabs()[0])
check("вкладка «Файлы» принимает файлы целиком, не только пунктирная зона",
      bool(files_tab.bind("<<Drop>>")))
check("по умолчанию выбраны оба формата: %s" % (app.selected_formats(),),
      app.selected_formats() == ("pdf", "docx"))

tabs = [app.notebook.tab(index, "text").strip() for index in app.notebook.tabs()]
check("есть вкладки «Файлы» и «Текст»: %s" % (tabs,), tabs == ["Файлы", "Текст"])

# вставка из буфера: подменяем чтение буфера, чтобы не трогать настоящий
app.root.clipboard_get = lambda: "# Из буфера\n\nФормула $x^2$ и текст.\n"
app.paste_from_clipboard()
root.update()
check("вставка из буфера попала в поле",
      "# Из буфера" in app.text_area.get("1.0", "end"))


def clipboard_without_text():
    raise module.tk.TclError("clipboard does not contain text")


app.root.clipboard_get = clipboard_without_text
app.paste_from_clipboard()
root.update()
check("буфер без текста объясняется по-русски",
      "нет текста" in app.log.get("1.0", "end"))

# имя файла для вставленного текста берётся из первого заголовка
sample = "# Отчёт по датчикам\n\nТекст с формулой $E=mc^2$.\n"
name = module.text_file_name(sample, PROJECT_DIR)
check("имя файла из заголовка: %s" % name.name, name.name == "Отчёт по датчикам.md")
check("без заголовка имя «Текст.md»", module.text_file_name("просто строка", PROJECT_DIR).name
      == "Текст.md")
check("папка для текста по умолчанию: %s" % module.default_text_dir(),
      module.default_text_dir().name == "PDFConv")

# пустой текст не запускает конвертацию, а объясняет, что делать
for variable in app.format_vars.values():
    variable.set(True)
app.text_area.delete("1.0", "end")
app.convert_text()
root.update()
check("пустой текст не запускает конвертацию", not app.running)
check("пустой текст объясняется в отчёте",
      "Текст пуст" in app.log.get("1.0", "end"))

for variable in app.format_vars.values():
    variable.set(False)
app.start_batch([PROJECT_DIR / "README.md"])  # файл не будет конвертирован: форматы сняты
root.update()
check("без выбранных форматов подсказка показана",
      "Отметьте хотя бы один формат" in app.log.get("1.0", "end"))
check("без выбранных форматов конвертация не запустилась", not app.running)

app.format_vars["docx"].set(True)
check("снят pdf — остался только docx: %s" % (app.selected_formats(),),
      app.selected_formats() == ("docx",))
app.format_vars["pdf"].set(True)
app.format_vars["docx"].set(False)
check("снят docx — остался только pdf: %s" % (app.selected_formats(),),
      app.selected_formats() == ("pdf",))

def wait_for_environment(timeout: float = 240.0) -> bool:
    """Проверка окружения идёт в потоке и включает пробную сборку LaTeX — ждём её."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        root.update()
        if not app.env_running:
            return True
        time.sleep(0.1)
    return False


# проверка окружения живёт в журнале: при запуске и по кнопке
# Стартует через after(60) после отрисовки, поэтому даём ей начаться, а не ловим
# гонку первых миллисекунд: на быстрой машине все проверки до этой успевают за 60 мс.
startup_deadline = time.time() + 5
while (not app.env_running and "Проверка окружения:" not in app.log.get("1.0", "end")
       and time.time() < startup_deadline):
    root.update()
    time.sleep(0.02)
startup_log = app.log.get("1.0", "end")
check("проверка окружения выполняется при старте",
      app.env_running or "Проверка окружения:" in startup_log)
wait_for_environment()
before = app.log.get("1.0", "end").count("Проверка окружения:")
app.refresh_environment()
wait_for_environment()
log_text = app.log.get("1.0", "end")
# строгое равенство было бы хрупким: проверка при запуске может отработать между замерами
check("кнопка «Проверить окружение снова» добавляет отчёт в журнал",
      log_text.count("Проверка окружения:") >= before + 1)
check("в журнале видно результат проверки",
      "Окружение в порядке" in log_text or "Нужно доустановить" in log_text)
check("проверка включает пробную сборку LaTeX", "Пробная сборка LaTeX" in log_text)
check("отдельной панели окружения больше нет", not hasattr(app, "env_body"))

# правый клик: журнал — «Копировать» и «Выделить всё»
menu_labels = [app.log_menu.entrycget(index, "label")
               for index in range(app.log_menu.index("end") + 1)]
check("в меню журнала есть «Копировать» и «Выделить всё»",
      menu_labels == ["Копировать", "Выделить всё"])
app.select_all_log()
root.update()
check("«Выделить всё» выделяет журнал", len(app.log.tag_ranges("sel")) > 0)
app.copy_log_selection()
copied = str(app.root.tk.call("clipboard", "get"))  # в обход подменённого clipboard_get
check("«Копировать» кладёт журнал в буфер обмена", len(copied.strip()) > 50)

# правый клик: поле текста — «Вставить» в позицию курсора
app.text_area.delete("1.0", "end")
app.text_area.insert("1.0", "начало ")
app.text_area.mark_set("insert", "end")
app.root.clipboard_get = lambda: "вставка"
app.text_menu.invoke(0)
root.update()
check("в меню поля текста есть «Вставить»",
      app.text_menu.entrycget(0, "label") == "Вставить")
check("правый клик вставляет из буфера в позицию курсора",
      app.text_area.get("1.0", "end").strip() == "начало вставка")

root.destroy()
print()
print("Все проверки окна пройдены." if not failures
      else "Не прошли проверки: %s" % ", ".join(failures))
sys.exit(1 if failures else 0)
"""Нарисовать icon.ico для программы: страница документа и стрелка преобразования.

Запуск: py -3 make_icon.py
Рисуем в 256 px и уменьшаем до остальных размеров — так края остаются гладкими.
"""
from pathlib import Path

from PIL import Image, ImageDraw

SIZE = 256
SIZES = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
BLUE = (11, 87, 208, 255)
BLUE_DARK = (8, 62, 150, 255)
PAGE = (255, 255, 255, 255)
LINE = (176, 190, 212, 255)
ACCENT = (255, 179, 0, 255)

image = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)

# фон: скруглённый квадрат с мягким затемнением снизу
draw.rounded_rectangle((8, 8, SIZE - 8, SIZE - 8), radius=52, fill=BLUE)
draw.rounded_rectangle((8, SIZE // 2, SIZE - 8, SIZE - 8), radius=52, fill=BLUE_DARK)
draw.rounded_rectangle((8, 8, SIZE - 8, SIZE - 8), radius=52, outline=(255, 255, 255, 40), width=4)

# лист документа с загнутым уголком
page_left, page_top, page_right, page_bottom = 58, 40, 158, 216
fold = 40
draw.polygon(
    [(page_left, page_top), (page_right - fold, page_top), (page_right, page_top + fold),
     (page_right, page_bottom), (page_left, page_bottom)],
    fill=PAGE,
)
draw.polygon(
    [(page_right - fold, page_top), (page_right - fold, page_top + fold),
     (page_right, page_top + fold)],
    fill=(210, 220, 235, 255),
)
for offset in range(3):  # строки текста
    top = page_top + 52 + offset * 26
    draw.rounded_rectangle((page_left + 16, top, page_right - 16, top + 10), radius=5, fill=LINE)

# стрелка преобразования: из документа наружу
shaft_top, shaft_bottom = 168, 196
draw.rounded_rectangle((126, shaft_top, 206, shaft_bottom), radius=14, fill=ACCENT)
draw.polygon([(196, 148), (238, 182), (196, 216)], fill=ACCENT)

res_dir = Path(__file__).resolve().parent.parent / "res"  # корень проекта / res
res_dir.mkdir(exist_ok=True)
icon_path = res_dir / "PDFConv.ico"
image.save(icon_path, sizes=SIZES)
print("готово: %s (%d байт, размеры: %s)" % (
    icon_path, icon_path.stat().st_size, ", ".join("%d" % width for width, _ in SIZES)))
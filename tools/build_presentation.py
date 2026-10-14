# -*- coding: utf-8 -*-
"""Презентация в техногенном стиле: фоны PIL + python-pptx + XML-анимации."""
import math, os, random
from PIL import Image, ImageDraw, ImageFilter, ImageFont

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BG = os.path.join("/tmp", "pres_bg")  # черновые фоны вне репозитория
IMG = os.path.join(ROOT, "ОВЧ-ЭКЗ", "Итоговые работы", "images")
OUT = os.path.join(ROOT, "ОВЧ-ЭКЗ", "Итоговые работы", "3952_Теддер_Д_ОхлаждениеКомпьютерныхСистем_Презентация.pptx")
os.makedirs(BG, exist_ok=True)

W, H = 1920, 1080
CYAN = (0, 229, 255)
MAGENTA = (255, 46, 136)
AMBER = (255, 176, 32)
LIME = (154, 255, 61)

FONT_M = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf"
FONT_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"

def font(path, size):
    return ImageFont.truetype(path, size)

def make_bg(idx, total, tag, accent=CYAN, kind="content"):
    img = Image.new("RGB", (W, H), (4, 6, 11))
    dr = ImageDraw.Draw(img, "RGBA")
    # вертикальный градиент
    for y in range(H):
        t = y / H
        c = (int(4 + 8 * t), int(6 + 14 * t), int(11 + 21 * t))
        dr.line([(0, y), (W, y)], fill=c + (255,))
    # сетка
    for x in range(0, W, 64):
        dr.line([(x, 0), (x, H)], fill=accent + (14,), width=1)
    for y in range(0, H, 64):
        dr.line([(0, y), (W, y)], fill=accent + (14,), width=1)
    # свечение
    glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    if kind == "title":
        gd.ellipse([W // 2 - 560, H // 2 - 420, W // 2 + 560, H // 2 + 420], outline=accent + (60,), width=3)
        gd.ellipse([W // 2 - 480, H // 2 - 350, W // 2 + 480, H // 2 + 350], outline=accent + (35,), width=2)
        gd.ellipse([W // 2 - 700, H // 2 - 520, W // 2 + 700, H // 2 + 520], fill=accent + (10,))
    else:
        gd.ellipse([-320, H - 560, 620, H + 380], fill=accent + (16,))
        gd.ellipse([W - 420, -320, W + 320, 420], fill=accent + (12,))
    glow = glow.filter(ImageFilter.GaussianBlur(30))
    img.paste(glow, (0, 0), glow)
    dr = ImageDraw.Draw(img, "RGBA")
    # сканлайны
    for y in range(0, H, 4):
        dr.line([(0, y), (W, y)], fill=(0, 0, 0, 26), width=1)
    # диагональные техно-штрихи справа
    rnd = random.Random(idx * 7 + 3)
    for i in range(14):
        x0 = W - rnd.randint(60, 420)
        y0 = rnd.randint(80, H - 80)
        ln = rnd.randint(40, 220)
        dr.line([(x0, y0), (x0 + ln, y0)], fill=accent + (rnd.randint(25, 70),), width=2)
        dr.rectangle([x0 - 6, y0 - 3, x0, y0 + 3], fill=accent + (90,))
    # уголки-скобки
    L = 56
    for (cx, cy, dx, dy) in [(28, 28, 1, 1), (W - 28, 28, -1, 1), (28, H - 28, 1, -1), (W - 28, H - 28, -1, -1)]:
        dr.line([(cx, cy), (cx + dx * L, cy)], fill=accent + (200,), width=4)
        dr.line([(cx, cy), (cx, cy + dy * L)], fill=accent + (200,), width=4)
    # верхняя и нижняя плашки
    dr.rectangle([0, 0, W, 6], fill=accent + (160,))
    dr.rectangle([0, H - 6, W, H], fill=accent + (90,))
    fm = font(FONT_M, 30)
    fs = font(FONT_M, 26)
    dr.text((64, 40), tag, font=fm, fill=accent + (230,))
    dr.text((W - 420, 42), "COOLING://ANALYSIS", font=fs, fill=(160, 190, 210, 200))
    dr.text((64, H - 62), "3952_ТЕДДЕР_Д // НовГУ_2026", font=fs, fill=(150, 175, 195, 190))
    num = f"{idx:02d} / {total:02d}"
    dr.text((W - 300, H - 66), num, font=fm, fill=accent + (235,))
    if kind != "title":
        dr.line([(64, 96), (W - 64, 96)], fill=accent + (70,), width=2)
    out = os.path.join(BG, f"slide{idx:02d}.png")
    img.save(out, quality=95)
    return out

# ------------------------------------------------------------------ слайды
SLIDES = [
    dict(tag="SYS://START", accent=CYAN, kind="title", title="АНАЛИЗ ТЕХНОЛОГИЙ ОХЛАЖДЕНИЯ\nКОМПЬЮТЕРНЫХ СИСТЕМ",
         sub="и основы их обслуживания",
         bullets=["Реферат по дисциплине «Техническое обслуживание и ремонт аппаратной части компьютерных систем и комплексов»",
                  "Выполнил: Теддер Даниил, группа 3952   Руководитель: Овчинников К.Е."],
         img=None),
    dict(tag="01 // АКТУАЛЬНОСТЬ", accent=CYAN, kind="content", title="Почему охлаждение решает всё",
         bullets=["Тепловой пакет флагманских ЦП и ГП достиг 125–350 Вт и продолжает расти",
                  "До 30 % отказов и нестабильности ПК связаны с деградацией охлаждения: пыль, высохшая термопаста, износ вентиляторов",
                  "Перегрев означает троттлинг, сбои, сокращение ресурса полупроводников",
                  "Обслуживание дешевле ремонта: 1–2 часа работ против замены компонентов",
                  "Цель реферата: проанализировать технологии охлаждения и сформировать регламент обслуживания"],
         img=None),
    dict(tag="02 // ФИЗИКА", accent=CYAN, kind="content", title="Физика процесса: цепь тепловых сопротивлений",
         bullets=["Три механизма переноса: теплопроводность, конвекция, излучение; лимитирует конвекция",
                  "Перепад температуры = тепловой поток × тепловое сопротивление участка",
                  "Пример: 150 Вт при сумме сопротивлений 0,38 К/Вт дают перегрев 57 °C над воздухом корпуса",
                  "Высохшая термопаста добавляет 0,10 К/Вт и уводит кристалл в троттлинг без замены кулера",
                  "Вывод: слабое звено цепи обесценивает всю конструкцию"],
         img=None),
    dict(tag="03 // ВОЗДУХ", accent=CYAN, kind="content", title="Воздушное охлаждение: рабочая лошадка отрасли",
         bullets=["Тепловые трубки: фазовый перенос тепла, эффективная проводимость в десятки раз выше меди",
                  "Башенные кулеры: до 250 Вт, две секции пластин, 6–8 трубок",
                  "Серверы: коридоры холодного и горячего воздуха, горячее резервирование вентиляторов",
                  "Плюсы: цена, надёжность, простота обслуживания   Минусы: габариты, шум, предел мощности"],
         img="noctua.jpg"),
    dict(tag="04 // ЖИДКОСТЬ", accent=CYAN, kind="content", title="Жидкостное охлаждение: моноблоки и контуры",
         bullets=["Моноблочная СЖО: герметичный контур, 250–350 Вт, 5+ лет без вмешательства",
                  "Пользовательский контур: водоблоки ЦП и ГП, 600 Вт и более, насос и резервуар",
                  "Жидкость переносит тепло циркуляцией, а не теплопроводностью: вода 0,6 Вт/м·К против 401 у меди",
                  "Плата за эффективность: цена, сложность, контроль герметичности и состояния жидкости"],
         img="aio.jpg"),
    dict(tag="05 // СПЕЦТЕХНОЛОГИИ", accent=MAGENTA, kind="content", title="Термоэлектричество, фаза, иммерсия",
         bullets=["Модуль Пельтье: температура ниже окружающей среды, но низкий КПД и риск конденсата",
                  "Компрессорные схемы: экстремальный разгон, защита от конденсата, нишевое применение",
                  "Иммерсионные ванны: платы в диэлектрике, PUE около 1,09 против 1,4–1,6 у воздушных залов",
                  "Пассивное охлаждение: 30–60 Вт без единого вентилятора, абсолютная тишина"],
         img="immersion.jpg"),
    dict(tag="06 // СРАВНЕНИЕ", accent=CYAN, kind="content", title="Сравнение технологий", table=dict(
         headers=["Технология", "Мощность", "Шум", "Обслуживание"],
         rows=[["Воздушное башенное", "до 250 Вт", "средний", "низкая сложность"],
               ["Моноблочная СЖО", "250–350 Вт", "низкий", "низкая сложность"],
               ["Пользовательский контур", "600 Вт и более", "низкий", "высокая сложность"],
               ["Термоэлектрическое", "до 150 Вт", "высокий", "средняя сложность"],
               ["Иммерсионное", "без предела узла", "минимальный", "средняя сложность"],
               ["Пассивное", "30–60 Вт", "нет", "минимальная"]]),
         bullets=["Универсального решения нет: выбор диктуют мощность, бюджет, шум и надёжность"],
         img=None),
    dict(tag="07 // ДЕГРАДАЦИЯ", accent=AMBER, kind="content", title="Почему охлаждение деградирует",
         bullets=["Пыль: войлок на фильтрах и радиаторе режет приток и теплоотдачу",
                  "Термопаста: за 2–4 года сохнет и трескается, сопротивление контакта растёт",
                  "Подшипники и насосы: износ смазки, шум, падение производительности",
                  "Химия контуров: окислы и биологические отложения засоряют микроканалы",
                  "Признаки: рост температур и оборотов, новый шум, видимые отложения"],
         img="dust2.jpg"),
    dict(tag="08 // КОНТРОЛЬ", accent=CYAN, kind="content", title="Нормы температур и средства контроля",
         table=dict(headers=["Компонент", "Покой", "Нагрузка", "Порог"],
                    rows=[["ЦП", "30–50 °C", "70–90 °C", "95–100 °C"],
                          ["ГП", "30–45 °C", "65–85 °C", "85–95 °C"],
                          ["SSD", "30–45 °C", "50–70 °C", "80 °C"]]),
         bullets=["HWiNFO64 и AIDA64: журналирование датчиков до и после работ",
                  "Prime95, OCCT, FurMark: идентичная нагрузка для честного сравнения",
                  "Кривые вентиляторов в прошивке: 30 % до 50 °C, максимум к 85 °C"],
         img=None),
    dict(tag="09 // РЕГЛАМЕНТ", accent=LIME, kind="content", title="Регламент обслуживания: восемь этапов",
         bullets=["1  Фиксация исходных температур и стресс-тест",
                  "2  Обесточивание и разборка, антистатика",
                  "3  Диагностика узлов и подшипников",
                  "4  Очистка от пыли, промывка фильтров",
                  "5  Замена термоинтерфейсов и прокладок",
                  "6  Обслуживание вентиляторов и контуров СЖО",
                  "7  Сборка и настройка кривых вращения",
                  "8  Контрольный прогон и карточка устройства"],
         img=None),
    dict(tag="10 // ЭФФЕКТ", accent=LIME, kind="content", title="Эффект обслуживания: пример на типовом ПК",
         table=dict(headers=["Показатель", "До", "После"],
                    rows=[["ЦП под нагрузкой", "91 °C", "72 °C"],
                          ["ГП под нагрузкой", "84 °C", "71 °C"],
                          ["Шум под нагрузкой", "41 дБА", "33 дБА"],
                          ["Троттлинг", "фиксировался", "устранён"]]),
         bullets=["Три года без обслуживания в доме с животными: процессор у порога и троттлинг",
                  "Стандартный регламент вернул 19 °C запаса и тишину без замены кулера"],
         img="paste.jpg"),
    dict(tag="11 // ПЕРИОДИЧНОСТЬ", accent=CYAN, kind="content", title="Периодичность и типовые ошибки",
         table=dict(headers=["Условия", "Чистка", "Термопаста"],
                    rows=[["Офис, дом без животных", "12 месяцев", "2–3 года"],
                          ["Животные, ремонт рядом", "6 месяцев", "1–2 года"],
                          ["Игры и рендеринг", "4–6 месяцев", "8–12 месяцев"]]),
         bullets=["Ошибка: продувка вентилятора без фиксации крыльчатки",
                  "Ошибка: избыток пасты и затяжка крепежа по одному винту",
                  "Ошибка: сборка без контрольного прогона с журналированием"],
         img=None),
    dict(tag="12 // БЕЗОПАСНОСТЬ", accent=AMBER, kind="content", title="Безопасность работ",
         bullets=["Электробезопасность: обесточивание, разряд ёмкостей, блок питания не вскрывать, приказ Минтруда 903н",
                  "Антистатика: браслет и коврик, выравнивание потенциалов перед касанием плат",
                  "Химия и воздух: очки, перчатки, проветривание, сжатый воздух короткими импульсами",
                  "СЖО: слив и заправка только на обесточенной системе, сутки контроля герметичности",
                  "Пожарная безопасность: углекислотный огнетушитель, вода под напряжением запрещена, телефон 101"],
         img=None),
    dict(tag="13 // ИТОГИ", accent=CYAN, kind="title", title="ВЫВОДЫ",
         bullets=["Проанализированы семь технологий охлаждения от воздушных до иммерсионных",
                  "Эффективность системы равна её слабому звену, поэтому обслуживание важнее цены кулера",
                  "Сформирован восьмиэтапный регламент с контролируемыми критериями результата",
                  "Регламент подтверждён примером: минус 19 °C и устранение троттлинга",
                  "Спасибо за внимание!"],
         img=None),
]

TOTAL = len(SLIDES)

# ------------------------------------------------------------------ сборка
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

EMU_IN = 914400

def add_text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    first = True
    for item in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = spacing
        for rspec in item.get("runs", [item]):
            r = p.add_run()
            r.text = rspec["t"]
            f = r.font
            f.name = rspec.get("font", "Consolas")
            f.size = Pt(rspec.get("size", 18))
            f.bold = rspec.get("bold", False)
            f.color.rgb = RGBColor.from_string(rspec.get("color", "E8F6FF"))
        if item.get("space_after"):
            p.space_after = Pt(item["space_after"])
    return tb

def add_table(slide, spec, x, y, w):
    rows = len(spec["rows"]) + 1
    cols = len(spec["headers"])
    gf = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(0.4 * rows))
    tbl = gf.table
    first = w * 0.34
    rest = (w - first) / (cols - 1)
    for ci, col in enumerate(tbl.columns):
        col.width = Inches(first if ci == 0 else rest)
    for c, htext in enumerate(spec["headers"]):
        cell = tbl.cell(0, c)
        cell.text = htext
        cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor.from_string("0E2A3A")
        p = cell.text_frame.paragraphs[0]
        p.font.name = "Consolas"; p.font.size = Pt(15); p.font.bold = True
        p.font.color.rgb = RGBColor.from_string("00E5FF")
    for r, row in enumerate(spec["rows"], start=1):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string("0A1622" if r % 2 else "0D1B2A")
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Consolas"; p.font.size = Pt(13)
            p.font.color.rgb = RGBColor.from_string("D7E9F5")
    return gf

def set_transition(slide, kind, idx):
    sld = slide._element
    trans = etree.SubElement(sld, qn("p:transition"))
    trans.set("spd", "med")
    if kind == "push":
        etree.SubElement(trans, qn("p:push")).set("dir", "l")
    elif kind == "cover":
        etree.SubElement(trans, qn("p:cover")).set("dir", "l")
    else:
        etree.SubElement(trans, qn("p:fade"))

def add_timing(slide, specs):
    """specs: list of (spid, trigger); trigger: 'auto' (с показом слайда) или 'click'."""
    sld = slide._element
    timing = etree.SubElement(sld, qn("p:timing"))
    tnlst = etree.SubElement(timing, qn("p:tnLst"))
    par0 = etree.SubElement(tnlst, qn("p:par"))
    ctn0 = etree.SubElement(par0, qn("p:cTn"))
    ctn0.set("id", "1"); ctn0.set("dur", "indefinite"); ctn0.set("restart", "never")
    child0 = etree.SubElement(ctn0, qn("p:childTnLst"))
    seq = etree.SubElement(child0, qn("p:seq"))
    seq.set("concurrent", "1"); seq.set("nextAc", "seek")
    ctn1 = etree.SubElement(seq, qn("p:cTn"))
    ctn1.set("id", "2"); ctn1.set("dur", "indefinite"); ctn1.set("nodeType", "tmRoot")
    child1 = etree.SubElement(ctn1, qn("p:childTnLst"))
    nid = 3
    cur_group = None
    for spid, trig in specs:
        if trig == "click" or cur_group is None:
            gpar = etree.SubElement(child1, qn("p:par"))
            gctn = etree.SubElement(gpar, qn("p:cTn"))
            gctn.set("id", str(nid)); nid += 1
            gctn.set("fill", "hold")
            gst = etree.SubElement(gctn, qn("p:stCondLst"))
            gcond = etree.SubElement(gst, qn("p:cond"))
            gcond.set("delay", "indefinite" if trig == "click" else "0")
            cur_group = etree.SubElement(gctn, qn("p:childTnLst"))
        mpar = etree.SubElement(cur_group, qn("p:par"))
        mctn = etree.SubElement(mpar, qn("p:cTn"))
        mctn.set("id", str(nid)); nid += 1
        mctn.set("fill", "hold")
        mst = etree.SubElement(mctn, qn("p:stCondLst"))
        mcond = etree.SubElement(mst, qn("p:cond")); mcond.set("delay", "0")
        mchild = etree.SubElement(mctn, qn("p:childTnLst"))
        epar = etree.SubElement(mchild, qn("p:par"))
        ectn = etree.SubElement(epar, qn("p:cTn"))
        ectn.set("id", str(nid)); nid += 1
        ectn.set("presetID", "10"); ectn.set("presetClass", "entr")
        ectn.set("presetSubtype", "0"); ectn.set("fill", "hold")
        ectn.set("nodeType", "clickEffect" if trig == "click" else "afterEffect")
        est = etree.SubElement(ectn, qn("p:stCondLst"))
        econ = etree.SubElement(est, qn("p:cond")); econ.set("delay", "0")
        echild = etree.SubElement(ectn, qn("p:childTnLst"))
        anim = etree.SubElement(echild, qn("p:animEffect"))
        anim.set("transition", "in"); anim.set("filter", "fade")
        cb = etree.SubElement(anim, qn("p:cBhvr"))
        ctn2 = etree.SubElement(cb, qn("p:cTn"))
        ctn2.set("id", str(nid)); nid += 1
        ctn2.set("dur", "450")
        tgt = etree.SubElement(cb, qn("p:tgtEl"))
        sp = etree.SubElement(tgt, qn("p:spTgt")); sp.set("spid", str(spid))
    prev = etree.SubElement(seq, qn("p:prevCondLst"))
    c = etree.SubElement(prev, qn("p:cond")); c.set("delay", "0"); c.set("evt", "onPrev")
    nxt = etree.SubElement(seq, qn("p:nextCondLst"))
    c = etree.SubElement(nxt, qn("p:cond")); c.set("delay", "0"); c.set("evt", "onNext")

shape_ids = []
for i, sd in enumerate(SLIDES, start=1):
    bgpath = make_bg(i, TOTAL, sd["tag"], sd["accent"], sd["kind"])
    slide = prs.slides.add_slide(blank)
    slide.shapes.add_picture(bgpath, 0, 0, width=prs.slide_width, height=prs.slide_height)
    ids = []
    if sd["kind"] == "title":
        if sd.get("sub"):
            t = add_text(slide, 1.0, 2.0, 11.33, 1.9,
                         [{"t": sd["title"], "size": 40, "bold": True, "color": "EAFBFF", "font": "Consolas"}],
                         align=PP_ALIGN.CENTER, spacing=1.05)
            ids.append(t.shape_id)
            sb = add_text(slide, 1.0, 3.95, 11.33, 0.6,
                          [{"t": sd["sub"], "size": 24, "color": "00E5FF", "font": "Consolas"}], align=PP_ALIGN.CENTER)
            ids.append(sb.shape_id)
            y = 4.85
            step, bsize = 0.66, 15
        else:
            t = add_text(slide, 1.0, 1.45, 11.33, 1.1,
                         [{"t": sd["title"], "size": 44, "bold": True, "color": "EAFBFF", "font": "Consolas"}],
                         align=PP_ALIGN.CENTER)
            ids.append(t.shape_id)
            y = 2.85
            step, bsize = 0.72, 17
        for b in sd["bullets"]:
            bb = add_text(slide, 1.4, y, 10.53, step,
                          [{"t": b, "size": bsize, "color": "B9D4E6", "font": "Segoe UI"}], align=PP_ALIGN.CENTER)
            ids.append(bb.shape_id)
            y += step
        specs = [(ids[0], "auto"), (ids[1], "auto")] + [(x, "click") for x in ids[2:]] if sd.get("sub") else \
                [(ids[0], "auto")] + [(x, "click") for x in ids[1:]]
    else:
        t = add_text(slide, 0.62, 0.62, 11.6, 0.9,
                     [{"t": sd["title"], "size": 30, "bold": True, "color": "EAFBFF", "font": "Consolas"}])
        auto_ids = [t.shape_id]
        click_ids = []
        y = 1.62
        has_table = "table" in sd
        has_img = sd.get("img")
        text_w = 7.35 if (has_img or has_table) else 11.9
        if has_table:
            tb = add_table(slide, sd["table"], 0.65, y + 0.1, 7.2)
            auto_ids.append(tb.shape_id)
            y += 0.42 * (len(sd["table"]["rows"]) + 1) + 0.35
        for b in sd["bullets"]:
            size = 16 if len(b) < 78 else 14
            bb = add_text(slide, 0.72, y, text_w - 0.1, 0.6,
                          [{"runs": [{"t": "> ", "size": size, "color": "00E5FF", "font": "Consolas", "bold": True},
                                     {"t": b, "size": size, "color": "D9ECF8", "font": "Segoe UI"}]}], spacing=1.0)
            click_ids.append(bb.shape_id)
            y += 0.62 if len(b) < 78 else 0.72
        if has_img:
            pth = os.path.join(IMG, sd["img"])
            from PIL import Image as PILImage
            iw, ih = PILImage.open(pth).size
            max_w, max_h = 4.5, 4.6
            w_in = min(max_w, (iw / ih) * max_h)
            h_in = w_in * ih / iw
            pic = slide.shapes.add_picture(pth, Inches(12.75 - w_in), Inches(3.55 - h_in / 2 + 0.6), width=Inches(w_in))
            auto_ids.append(pic.shape_id)
        specs = [(x, "auto") for x in auto_ids] + [(x, "click") for x in click_ids]
    set_transition(slide, ["push", "fade", "cover"][i % 3], i)
    add_timing(slide, specs)

prs.core_properties.title = "Анализ технологий охлаждения компьютерных систем"
prs.core_properties.author = "Теддер Даниил"
prs.save(OUT)
print("Сохранено:", OUT, "слайдов:", TOTAL)

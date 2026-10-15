# -*- coding: utf-8 -*-
"""
Улучшенная презентация в техногенном стиле (Версия 2):
- Раздел сравнения с прямым ответом на вопрос «Какой способ охлаждения всё-таки лучше?»:
  * Сравнительная таблица слева
  * Выделенный интерактивный блок с вердиктом и сценариями выбора справа
- Продвинутые сложные анимации (Fly, Wipe, Zoom, Fade, Blinds) с каскадными таймингами
- Богатый фактаж из реферата (TDP, R_th, формулы, термопасты, нормы ГОСТ и охраны труда)
- Иллюстрации из папки images
"""
import math, os, random
from PIL import Image, ImageDraw, ImageFilter, ImageFont

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if not os.path.exists(os.path.join(ROOT, "ОВЧ-ЭКЗ")):
    ROOT = "/home/user/Referat-Ovechka"

BG = os.path.join("/tmp", "pres_bg_upgraded2")
IMG = os.path.join(ROOT, "ОВЧ-ЭКЗ", "Итоговые работы", "images")
OUT = os.path.join(ROOT, "ОВЧ-ЭКЗ", "Итоговые работы", "3952_Теддер_Д_ОхлаждениеКомпьютерныхСистем_Презентация.pptx")
OUT2 = os.path.join(ROOT, "ОВЧ-ЭКЗ", "Итоговые работы2", "3952_Теддер_Д_ОхлаждениеКомпьютерныхСистем_Презентация2.pptx")
os.makedirs(BG, exist_ok=True)
os.makedirs(os.path.dirname(OUT2), exist_ok=True)

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

# ------------------------------------------------------------------ определение слайдов
SLIDES = [
    dict(tag="SYS://START", accent=CYAN, kind="title",
         title="АНАЛИЗ ТЕХНОЛОГИЙ ОХЛАЖДЕНИЯ\nКОМПЬЮТЕРНЫХ СИСТЕМ",
         sub="и основы их технического обслуживания",
         bullets=["Реферат по дисциплине «Техническое обслуживание и ремонт аппаратной части компьютерных систем и комплексов»",
                  "Выполнил: Теддер Даниил, группа 3952   |   Руководитель: Овчинников К.Е."],
         img=None,
         title_anim="zoom",
         sub_anim="fade",
         bullet_anim="fly_bottom"),

    dict(tag="01 // АКТУАЛЬНОСТЬ", accent=CYAN, kind="content",
         title="Почему охлаждение решает всё: тепловой вызов",
         bullets=["Тепловой пакет флагманских ЦП и ГП достиг 125–350 Вт, плотность теплового потока превышает 100 Вт/см²",
                  "До 30 % аппаратных сбоев вызваны деградацией охлаждения: пыль, деградация термопасты, износ подшипников",
                  "Перегрев приводит к троттлингу частот, аварийным перезагрузкам и ускоренной электромиграции полупроводников",
                  "Своевременный регламент обслуживания (1–2 часа) в 10–20 раз дешевле внепланового ремонта и замены чипов",
                  "Цель исследования: сравнительный анализ технологий охлаждения и формирование строгого регламента ТО"],
         img="thermo.jpg",
         title_anim="fly_left",
         bullet_anim="wipe_right",
         img_anim="zoom"),

    dict(tag="02 // ФИЗИКА", accent=CYAN, kind="content",
         title="Физика процесса: закон Фурье и цепь тепловых сопротивлений",
         bullets=["Три механизма теплопереноса: теплопроводность, конвекция и излучение; узкое горлышко — конвекция в радиаторе",
                  "Формула теплового баланса: ΔT = P × ΣR_th (перепад температур = мощность × сумма сопротивлений)",
                  "Пример: 150 Вт при сумме R_th = 0,38 К/Вт дают перегрев кристалла 57 °C над температурой воздуха в корпусе",
                  "Деградация термоинтерфейса добавляет 0,10–0,15 К/Вт, вызывая критический перегрев даже при мощном кулере",
                  "Главный инженерный принцип: эффективность охлаждения лимитируется самым слабым звеном тепловой цепи"],
         img="heatpipe.png",
         title_anim="fly_left",
         bullet_anim="fly_left",
         img_anim="wipe_right"),

    dict(tag="03 // ВОЗДУХ", accent=CYAN, kind="content",
         title="Воздушное охлаждение: тепловые трубки и башни",
         bullets=["Тепловые трубки: испарительно-конденсационный цикл, эквивалентная теплопроводность до десятков тысяч Вт/(м·К)",
                  "Современные башенные кулеры: двухсекционные радиаторы, 6–8 трубок диаметром 6 мм, рассеивание до 250 Вт",
                  "Серверные решения: вентиляторные модули горячей замены (до 15 000 об/мин), организация горячих и холодных коридоров",
                  "Достоинства: максимальная надёжность, отсутствие риска протечек, низкая стоимость, простота регламента"],
         img="noctua.jpg",
         title_anim="wipe_left",
         bullet_anim="fly_left",
         img_anim="zoom"),

    dict(tag="04 // ЖИДКОСТЬ", accent=CYAN, kind="content",
         title="Жидкостное охлаждение: моноблоки AIO и кастомные контуры",
         bullets=["Моноблочные СЖО (AIO): необслуживаемый герметичный контур, радиаторы 240–360 мм, рассеивание 250–350 Вт",
                  "Кастомные контуры: раздельные водоблоки ЦП и ГП, помпа D5/DDC, резервуар, медные радиаторы — отвод от 600 Вт",
                  "Высокая удельная теплоёмкость воды (4184 Дж/(кг·К)) сглаживает пиковые тепловые удары на кристаллах",
                  "Цена эффективности: сложность монтажа, мониторинг герметичности, риск коррозии и требование ежегодного ТО"],
         img="aio.jpg",
         title_anim="fly_left",
         bullet_anim="wipe_right",
         img_anim="fly_right"),

    dict(tag="05 // КАСТОМ СЖО", accent=CYAN, kind="content",
         title="Архитектура контурного жидкостного охлаждения",
         bullets=["Ключевые узлы: водоблоки микроканального типа (каналы 0,1–0,2 мм), помпа высокой подачи, радиаторы с медными ламелями",
                  "Особое требование к совместимости материалов: строгий запрет на контакт меди и алюминия (гальваническая пара)",
                  "Теплоноситель: бидистиллированная вода со специализированными антикоррозийными и биоцидными присадками",
                  "Обслуживание: опрессовка контура давлением 0,5 бар перед заправкой и контроль прозрачности хладагента"],
         img="custom.jpg",
         title_anim="fly_bottom",
         bullet_anim="fly_left",
         img_anim="zoom"),

    dict(tag="06 // СПЕЦТЕХНОЛОГИИ", accent=MAGENTA, kind="content",
         title="Специальные технологии: элементы Пельтье, фаза и иммерсия",
         bullets=["Термоэлектрический модуль Пельтье: охлаждение ниже среды, но низкий КПД (COP < 0,5) и угроза точки росы",
                  "Компрессорное фазовое охлаждение: минусовые температуры под экстремальный разгон, изоляция от конденсата",
                  "Иммерсионные ванны (погружение в диэлектрические жидкости 3M Novec/синтетические масла): PUE центров данных до 1,09",
                  "Пассивные радиаторные системы: 30–65 Вт, абсолютная акустическая тишина, нулевой износ движущихся частей"],
         img="peltier.jpg",
         title_anim="fly_left",
         bullet_anim="wipe_right",
         img_anim="zoom"),

    dict(tag="07 // СРАВНЕНИЕ", accent=CYAN, kind="content",
         title="Сравнительный анализ: какой способ охлаждения лучше?",
         table=dict(
             headers=["Технология", "Предел TDP", "Шум", "Сложность ТО", "Надёжность"],
             rows=[["Башенный кулер", "до 250 Вт", "средний", "минимальная", "максимальная"],
                   ["Моноблок AIO", "250–350 Вт", "низкий", "низкая", "высокая (3–5 лет)"],
                   ["Кастомная СЖО", "600+ Вт", "минимальный", "очень высокая", "требует ТО"],
                   ["Модуль Пельтье", "до 150 Вт", "высокий", "средняя", "риск конденсата"],
                   ["Иммерсионное", "не ограничен", "минимальный", "средняя", "высокая (ЦОД)"],
                   ["Пассивное", "30–60 Вт", "бесшумно", "минимальная", "абсолютная"]]),
         side_panel=dict(
             title="ВЕРДИКТ: ЧТО ЖЕ ВСЁ-ТАКИ ЛУЧШЕ?",
             items=[
                 {"tag": "★ АБСОЛЮТНЫЙ ЛИДЕР ДЛЯ 90% ПК:", "text": "Воздушный башенный кулер (6–8 трубок). Лучшее соотношение цена/надёжность/эффективность: не может протечь, служит 10+ лет."},
                 {"tag": "★ ЛУЧШИЙ ВЫБОР ДЛЯ ГОРЯЧИХ ЦП (>250 Вт):", "text": "Моноблочная СЖО (AIO 280/360 мм). Идеальна для компактных корпусов и тяжёлого рендеринга без шума."},
                 {"tag": "★ ЭТАЛОН ДЛЯ ЭНТУЗИАСТОВ (ЦП + ГП):", "text": "Кастомный контур. Отводит свыше 600 Вт в абсолютной тишине, но требует регулярного ТО."},
                 {"tag": "★ ПРОРЫВ ДЛЯ СЕРВЕРОВ И ЦОД:", "text": "Иммерсионные ванны. Достигают PUE = 1,09 и рекордной плотности вычислений."}
             ]
         ),
         bullets=["Итог: единого идеала нет, но для ПК стандарт надёжности — воздух, а для пиковых нагрузок — СЖО"],
         img=None,
         title_anim="fly_left"),

    dict(tag="08 // ДЕГРАДАЦИЯ", accent=AMBER, kind="content",
         title="Факторы деградации систем охлаждения со временем",
         bullets=["Пылевой войлок на радиаторах и фильтрах: многократное снижение расхода воздуха и коэффициента конвекции",
                  "Деградация термоинтерфейса: испарение кремнийорганических масел, растрескивание матрицы через 2–3 года",
                  "Износ гидродинамических (FDB) и шариковых подшипников: выработка смазки, вибрация крыльчатки, шум",
                  "Засорение микроканалов СЖО: выпадение солевых осадков, биопленка и гальванический шлам в контуре",
                  "Симптомы: вентиляторы на 100% оборотов, нагрев ЦП выше 85 °C в простых задачах, просадки FPS и троттлинг"],
         img="dust2.jpg",
         title_anim="wipe_left",
         bullet_anim="fly_left",
         img_anim="zoom"),

    dict(tag="09 // МОНИТОРИНГ", accent=CYAN, kind="content",
         title="Инструментальный контроль и мониторинг состояния",
         table=dict(
             headers=["Компонент", "В покое", "Под нагрузкой", "Критический порог"],
             rows=[["Центральный процессор (CPU)", "30–45 °C", "65–85 °C", "95–100 °C (T_jmax)"],
                   ["Видеокарта (GPU Core / HotSpot)", "30–40 °C", "60–78 °C", "85 °C / 105 °C"],
                   ["Твердотельный накопитель NVMe", "30–42 °C", "50–68 °C", "75–80 °C (троттлинг)"],
                   ["Цепи питания VRM материнки", "35–50 °C", "65–85 °C", "105–115 °C"]]),
         bullets=["Программные комплексы HWiNFO64 и AIDA64: считывание датчиков температур, мощности и оборотов вентиляторов",
                  "Стресс-тестирование в Prime95 (Small FFTs), OCCT и FurMark для верификации теплового баланса до и после ТО",
                  "Настройка профилей вращения (Fan Curve): тихий режим до 50 °C, линейный рост до 75 °C, 100% при 85 °C"],
         img="hwinfo.jpg",
         title_anim="fly_left",
         bullet_anim="fly_left",
         img_anim="fade"),

    dict(tag="10 // РЕГЛАМЕНТ", accent=LIME, kind="content",
         title="Пошаговый регламент технического обслуживания (8 этапов)",
         bullets=["1  Входная диагностика: стресс-тест в ОС, фиксация температур T_max, оборотов RPM и шума",
                  "2  Обесточивание: отключение кабеля 220 В, разряд емкостей БП кнопкой Power, антистатический браслет",
                  "3  Разборка и аудит: визуальный осмотр радиаторов, лопастей, шлангов СЖО, проверка люфта подшипников",
                  "4  Очистка: продувка радиаторов сжатым воздухом, промывка пылевых сеток шасси нейтральным мыльным раствором",
                  "5  Замена термоинтерфейсов: удаление засохшей пасты изопропанолом 99,6%, нанесение свежей пасты (MX-4 / Kryonaut)",
                  "6  Обслуживание приводов: очистка вентиляторов, проверка помпы СЖО, промывка контура при необходимости",
                  "7  Монтаж и сборка: равномерная перекрестная затяжка винтов кулера, исключающая перекос кристалла",
                  "8  Выходной контроль: 30-минутный стресс-тест с логированием температур и оформлением паспорта устройства"],
         img=None,
         title_anim="fly_bottom",
         bullet_anim="fly_left"),

    dict(tag="11 // ЭФФЕКТ", accent=LIME, kind="content",
         title="Практический эффект обслуживания на типовом ПК",
         table=dict(
             headers=["Параметр системы", "До регламента ТО", "После регламента ТО", "Разница / Выигрыш"],
             rows=[["Температура CPU под нагрузкой", "91 °C (троттлинг)", "72 °C", "-19 °C"],
                   ["Температура GPU под нагрузкой", "84 °C", "71 °C", "-13 °C"],
                   ["Акустический шум (на 1 м)", "41 дБА (сильный гул)", "33 дБА (комфортный фон)", "-8 дБА"],
                   ["Тактовая частота всех ядер", "3,8 ГГц (сброс частот)", "4,4 ГГц (номинальный Boost)", "+600 МГц (+16 %)"],
                   ["Статус термотроттлинга", "Активен (Throttling YES)", "Отсутствует (Throttling NO)", "100 % стабильность"]]),
         bullets=["Анализ кейса: системный блок после 3 лет домашней эксплуатации без чистки радиаторов и смены пасты",
                  "Регламент ТО полностью восстановил заводские тепловые параметры и тишину без покупки новых комплектующих"],
         img="paste.jpg",
         title_anim="fly_left",
         bullet_anim="wipe_right",
         img_anim="zoom"),

    dict(tag="12 // ПЕРИОДИЧНОСТЬ", accent=CYAN, kind="content",
         title="Периодичность обслуживания и типичные ошибки",
         table=dict(
             headers=["Условия эксплуатации ПК", "Очистка от пыли", "Замена термопасты", "Контроль СЖО"],
             rows=[["Офис, чистое жилое помещение", "раз в 12 мес.", "каждые 2–3 года", "раз в 12 мес."],
                   ["Дом с домашними животными", "раз в 6 мес.", "каждые 1,5–2 года", "раз в 6 мес."],
                   ["Мастерские, запыленные помещения", "раз в 3 мес.", "раз в 12 мес.", "каждые 3–6 мес."],
                   ["Игровые клубы / Рендер-станции", "раз в 4–6 мес.", "раз в 12 мес.", "раз в 6 мес."]]),
         bullets=["Грубая ошибка: раскрутка вентиляторов струей сжатого воздуха без фиксации крыльчатки (ЭДС генератора сжигает плату)",
                  "Грубая ошибка: избыток или дефицит термопасты, неравномерная затяжка винтов с перекосом прижима",
                  "Грубая ошибка: смешивание медных и алюминиевых компонентов в одном контуре жидкостного охлаждения"],
         img="fans.jpg",
         title_anim="wipe_left",
         bullet_anim="fly_left",
         img_anim="zoom"),

    dict(tag="13 // БЕЗОПАСНОСТЬ", accent=AMBER, kind="content",
         title="Охрана труда и техника безопасности при проведении ТО",
         bullets=["Электробезопасность (приказ Минтруда РФ № 903н): обязательное физическое отключение от сети, разряд конденсаторов БП",
                  "Категорический запрет: не вскрывать корпус блока питания лицам без соответствующей группы допуска по электробезопасности",
                  "Антистатическая защита (ГОСТ IEC 61340): использование антистатического браслета и коврика, снятие заряда перед касанием плат",
                  "Химическая безопасность: работа с изопропиловым спиртом в проветриваемом помещении в перчатках, защита глаз при продувке сжатым газом",
                  "Пожарная безопасность (ППР РФ № 1479): оснащение рабочего места углекислотным огнетушителем (ОУ), запрет тушения электроники водой"],
         img=None,
         title_anim="fly_left",
         bullet_anim="fly_left"),

    dict(tag="14 // ИТОГИ", accent=CYAN, kind="title",
         title="ЗАКЛЮЧЕНИЕ И ВЫВОДЫ",
         sub="Результаты теоретического анализа и регламента обслуживания",
         bullets=["Проведен сравнительный анализ ключевых технологий: воздушных, жидкостных (AIO/кастом), термоэлектрических и иммерсионных",
                  "Доказано: надежность и тепловой режим системы определяются R_th самого слабого звена цепи — термоинтерфейса и чистоты ламелей",
                  "Разработан и апробирован комплексный 8-этапный регламент ТО с объективным инструментальным контролем параметров",
                  "Практический эффект регламента подтвержден: снижение температуры CPU на 19 °C, устранение троттлинга и снижение шума на 8 дБА",
                  "Спасибо за внимание! Готов ответить на ваши вопросы."],
         img=None,
         title_anim="zoom",
         sub_anim="fade",
         bullet_anim="fly_bottom"),
]

TOTAL = len(SLIDES)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

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
    gf = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(0.38 * rows))
    tbl = gf.table
    first = w * (0.30 if cols == 5 else (0.28 if cols == 4 else 0.35))
    rest = (w - first) / (cols - 1)
    for ci, col in enumerate(tbl.columns):
        col.width = Inches(first if ci == 0 else rest)
    for c, htext in enumerate(spec["headers"]):
        cell = tbl.cell(0, c)
        cell.text = htext
        cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor.from_string("0E2A3A")
        p = cell.text_frame.paragraphs[0]
        p.font.name = "Consolas"; p.font.size = Pt(13); p.font.bold = True
        p.font.color.rgb = RGBColor.from_string("00E5FF")
    for r, row in enumerate(spec["rows"], start=1):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string("0A1622" if r % 2 else "0D1B2A")
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Consolas"; p.font.size = Pt(11.5)
            p.font.color.rgb = RGBColor.from_string("D7E9F5")
    return gf

def set_transition(slide, kind):
    sld = slide._element
    trans = etree.SubElement(sld, qn("p:transition"))
    trans.set("spd", "med")
    if kind == "push":
        etree.SubElement(trans, qn("p:push")).set("dir", "l")
    elif kind == "cover":
        etree.SubElement(trans, qn("p:cover")).set("dir", "l")
    elif kind == "wipe":
        etree.SubElement(trans, qn("p:wipe")).set("dir", "l")
    else:
        etree.SubElement(trans, qn("p:fade"))

def build_advanced_anim(slide, anim_specs):
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
    for spec in anim_specs:
        spid = spec["spid"]
        trig = spec.get("trig", "click")
        effect = spec.get("effect", "fade")
        dur = str(spec.get("dur", 450))
        delay = str(spec.get("delay", 0))

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
        mcond = etree.SubElement(mst, qn("p:cond")); mcond.set("delay", delay)
        mchild = etree.SubElement(mctn, qn("p:childTnLst"))

        epar = etree.SubElement(mchild, qn("p:par"))
        ectn = etree.SubElement(epar, qn("p:cTn"))
        ectn.set("id", str(nid)); nid += 1
        ectn.set("presetClass", "entr")
        ectn.set("fill", "hold")
        ectn.set("nodeType", "clickEffect" if trig == "click" else "afterEffect")
        est = etree.SubElement(ectn, qn("p:stCondLst"))
        econ = etree.SubElement(est, qn("p:cond")); econ.set("delay", "0")
        echild = etree.SubElement(ectn, qn("p:childTnLst"))

        if effect == "zoom":
            ectn.set("presetID", "23")
            ectn.set("presetSubtype", "0")
            anim = etree.SubElement(echild, qn("p:animScale"))
            cb = etree.SubElement(anim, qn("p:cBhvr"))
            ctn2 = etree.SubElement(cb, qn("p:cTn"))
            ctn2.set("id", str(nid)); nid += 1
            ctn2.set("dur", dur)
            tgt = etree.SubElement(cb, qn("p:tgtEl"))
            sp = etree.SubElement(tgt, qn("p:spTgt")); sp.set("spid", str(spid))
            anim_fade = etree.SubElement(echild, qn("p:animEffect"))
            anim_fade.set("transition", "in"); anim_fade.set("filter", "fade")
            cb2 = etree.SubElement(anim_fade, qn("p:cBhvr"))
            ctn3 = etree.SubElement(cb2, qn("p:cTn"))
            ctn3.set("id", str(nid)); nid += 1
            ctn3.set("dur", dur)
            tgt2 = etree.SubElement(cb2, qn("p:tgtEl"))
            sp2 = etree.SubElement(tgt2, qn("p:spTgt")); sp2.set("spid", str(spid))
        elif effect.startswith("fly_"):
            ectn.set("presetID", "2")
            anim = etree.SubElement(echild, qn("p:animEffect"))
            anim.set("transition", "in")
            if effect == "fly_bottom":
                anim.set("filter", "fly(down)")
                ectn.set("presetSubtype", "4")
            elif effect == "fly_left":
                anim.set("filter", "fly(left)")
                ectn.set("presetSubtype", "1")
            elif effect == "fly_right":
                anim.set("filter", "fly(right)")
                ectn.set("presetSubtype", "8")
            cb = etree.SubElement(anim, qn("p:cBhvr"))
            ctn2 = etree.SubElement(cb, qn("p:cTn"))
            ctn2.set("id", str(nid)); nid += 1
            ctn2.set("dur", dur)
            tgt = etree.SubElement(cb, qn("p:tgtEl"))
            sp = etree.SubElement(tgt, qn("p:spTgt")); sp.set("spid", str(spid))
        elif effect == "wipe_left":
            ectn.set("presetID", "22")
            ectn.set("presetSubtype", "1")
            anim = etree.SubElement(echild, qn("p:animEffect"))
            anim.set("transition", "in"); anim.set("filter", "wipe(left)")
            cb = etree.SubElement(anim, qn("p:cBhvr"))
            ctn2 = etree.SubElement(cb, qn("p:cTn"))
            ctn2.set("id", str(nid)); nid += 1
            ctn2.set("dur", dur)
            tgt = etree.SubElement(cb, qn("p:tgtEl"))
            sp = etree.SubElement(tgt, qn("p:spTgt")); sp.set("spid", str(spid))
        elif effect == "wipe_right":
            ectn.set("presetID", "22")
            ectn.set("presetSubtype", "8")
            anim = etree.SubElement(echild, qn("p:animEffect"))
            anim.set("transition", "in"); anim.set("filter", "wipe(right)")
            cb = etree.SubElement(anim, qn("p:cBhvr"))
            ctn2 = etree.SubElement(cb, qn("p:cTn"))
            ctn2.set("id", str(nid)); nid += 1
            ctn2.set("dur", dur)
            tgt = etree.SubElement(cb, qn("p:tgtEl"))
            sp = etree.SubElement(tgt, qn("p:spTgt")); sp.set("spid", str(spid))
        elif effect == "blinds":
            ectn.set("presetID", "7")
            ectn.set("presetSubtype", "1")
            anim = etree.SubElement(echild, qn("p:animEffect"))
            anim.set("transition", "in"); anim.set("filter", "blinds(horizontal)")
            cb = etree.SubElement(anim, qn("p:cBhvr"))
            ctn2 = etree.SubElement(cb, qn("p:cTn"))
            ctn2.set("id", str(nid)); nid += 1
            ctn2.set("dur", dur)
            tgt = etree.SubElement(cb, qn("p:tgtEl"))
            sp = etree.SubElement(tgt, qn("p:spTgt")); sp.set("spid", str(spid))
        elif effect == "wheel":
            ectn.set("presetID", "21")
            ectn.set("presetSubtype", "4")
            anim = etree.SubElement(echild, qn("p:animEffect"))
            anim.set("transition", "in"); anim.set("filter", "wheel(4)")
            cb = etree.SubElement(anim, qn("p:cBhvr"))
            ctn2 = etree.SubElement(cb, qn("p:cTn"))
            ctn2.set("id", str(nid)); nid += 1
            ctn2.set("dur", dur)
            tgt = etree.SubElement(cb, qn("p:tgtEl"))
            sp = etree.SubElement(tgt, qn("p:spTgt")); sp.set("spid", str(spid))
        else: # fade
            ectn.set("presetID", "10")
            ectn.set("presetSubtype", "0")
            anim = etree.SubElement(echild, qn("p:animEffect"))
            anim.set("transition", "in"); anim.set("filter", "fade")
            cb = etree.SubElement(anim, qn("p:cBhvr"))
            ctn2 = etree.SubElement(cb, qn("p:cTn"))
            ctn2.set("id", str(nid)); nid += 1
            ctn2.set("dur", dur)
            tgt = etree.SubElement(cb, qn("p:tgtEl"))
            sp = etree.SubElement(tgt, qn("p:spTgt")); sp.set("spid", str(spid))

    prev = etree.SubElement(seq, qn("p:prevCondLst"))
    c = etree.SubElement(prev, qn("p:cond")); c.set("delay", "0"); c.set("evt", "onPrev")
    nxt = etree.SubElement(seq, qn("p:nextCondLst"))
    c = etree.SubElement(nxt, qn("p:cond")); c.set("delay", "0"); c.set("evt", "onNext")

TRANSITIONS = ["push", "cover", "fade", "wipe"]

for i, sd in enumerate(SLIDES, start=1):
    bgpath = make_bg(i, TOTAL, sd["tag"], sd["accent"], sd["kind"])
    slide = prs.slides.add_slide(blank)
    slide.shapes.add_picture(bgpath, 0, 0, width=prs.slide_width, height=prs.slide_height)

    anim_specs = []

    if sd["kind"] == "title":
        t_effect = sd.get("title_anim", "zoom")
        sub_effect = sd.get("sub_anim", "fade")
        b_effect = sd.get("bullet_anim", "fly_bottom")

        if sd.get("sub"):
            t = add_text(slide, 1.0, 1.85, 11.33, 1.9,
                         [{"t": sd["title"], "size": 38, "bold": True, "color": "EAFBFF", "font": "Consolas"}],
                         align=PP_ALIGN.CENTER, spacing=1.05)
            anim_specs.append({"spid": t.shape_id, "trig": "auto", "effect": t_effect, "dur": 550, "delay": 0})

            sb = add_text(slide, 1.0, 3.85, 11.33, 0.6,
                          [{"t": sd["sub"], "size": 23, "color": "00E5FF", "font": "Consolas"}], align=PP_ALIGN.CENTER)
            anim_specs.append({"spid": sb.shape_id, "trig": "auto", "effect": sub_effect, "dur": 450, "delay": 200})

            y = 4.80
            step, bsize = 0.64, 15
        else:
            t = add_text(slide, 1.0, 1.35, 11.33, 1.1,
                         [{"t": sd["title"], "size": 42, "bold": True, "color": "EAFBFF", "font": "Consolas"}],
                         align=PP_ALIGN.CENTER)
            anim_specs.append({"spid": t.shape_id, "trig": "auto", "effect": t_effect, "dur": 550, "delay": 0})
            y = 2.65
            step, bsize = 0.70, 16

        for b_idx, b in enumerate(sd["bullets"]):
            bb = add_text(slide, 1.4, y, 10.53, step,
                          [{"t": b, "size": bsize, "color": "B9D4E6", "font": "Segoe UI"}], align=PP_ALIGN.CENTER)
            if sd.get("sub"):
                anim_specs.append({"spid": bb.shape_id, "trig": "auto", "effect": b_effect, "dur": 400, "delay": 400 + b_idx * 250})
            else:
                anim_specs.append({"spid": bb.shape_id, "trig": "click", "effect": b_effect, "dur": 450})
            y += step

    else:
        # Контентный слайд
        t_effect = sd.get("title_anim", "fly_left")
        b_effect = sd.get("bullet_anim", "wipe_right")
        img_effect = sd.get("img_anim", "zoom")

        t = add_text(slide, 0.62, 0.62, 11.6, 0.9,
                     [{"t": sd["title"], "size": 28, "bold": True, "color": "EAFBFF", "font": "Consolas"}])
        anim_specs.append({"spid": t.shape_id, "trig": "auto", "effect": t_effect, "dur": 450, "delay": 0})

        has_table = "table" in sd
        has_side_panel = "side_panel" in sd
        has_img = sd.get("img")

        if has_side_panel:
            # Слайд сравнения: таблица слева, блок вердикта справа
            y = 1.58
            tb = add_table(slide, sd["table"], 0.65, y + 0.05, 7.3)
            anim_specs.append({"spid": tb.shape_id, "trig": "auto", "effect": "wipe_right", "dur": 500, "delay": 150})

            # Правый блок: заголовок вердикта + 4 карточки
            sp_data = sd["side_panel"]
            sp_hdr = add_text(slide, 8.20, y + 0.05, 4.50, 0.45,
                              [{"t": sp_data["title"], "size": 15, "bold": True, "color": "00E5FF", "font": "Consolas"}])
            anim_specs.append({"spid": sp_hdr.shape_id, "trig": "auto", "effect": "fade", "dur": 450, "delay": 250})

            py = y + 0.52
            for itm in sp_data["items"]:
                card_runs = [
                    {"runs": [{"t": itm["tag"] + "\n", "size": 12.5, "bold": True, "color": "9AFF3D", "font": "Consolas"},
                              {"t": itm["text"], "size": 12, "color": "E2F1FB", "font": "Segoe UI"}]}
                ]
                card = add_text(slide, 8.20, py, 4.50, 1.05, card_runs, spacing=1.0)
                anim_specs.append({"spid": card.shape_id, "trig": "click", "effect": "fly_right", "dur": 400})
                py += 1.10

            # Итоговый вывод внизу под таблицей
            by = y + 0.40 * (len(sd["table"]["rows"]) + 1) + 0.35
            for b in sd["bullets"]:
                bb = add_text(slide, 0.72, by, 7.2, 0.65,
                              [{"runs": [{"t": "⚡ ИТОГ: ", "size": 13.5, "color": "FFB020", "font": "Consolas", "bold": True},
                                         {"t": b, "size": 13.5, "color": "D9ECF8", "font": "Segoe UI"}]}], spacing=1.0)
                anim_specs.append({"spid": bb.shape_id, "trig": "click", "effect": "wipe_right", "dur": 400})
                by += 0.65

        else:
            text_w = 7.35 if (has_img or has_table) else 11.9
            y = 1.58

            if has_table:
                tb = add_table(slide, sd["table"], 0.65, y + 0.05, 7.2)
                anim_specs.append({"spid": tb.shape_id, "trig": "auto", "effect": "wipe_right", "dur": 500, "delay": 150})
                y += 0.40 * (len(sd["table"]["rows"]) + 1) + 0.30

            for b in sd["bullets"]:
                size = 15 if len(b) < 80 else 13.5
                bb = add_text(slide, 0.72, y, text_w - 0.1, 0.58,
                              [{"runs": [{"t": "> ", "size": size, "color": "00E5FF", "font": "Consolas", "bold": True},
                                         {"t": b, "size": size, "color": "D9ECF8", "font": "Segoe UI"}]}], spacing=1.0)
                anim_specs.append({"spid": bb.shape_id, "trig": "click", "effect": b_effect, "dur": 400})
                y += 0.58 if len(b) < 80 else 0.68

            if has_img:
                pth = os.path.join(IMG, sd["img"])
                from PIL import Image as PILImage
                iw, ih = PILImage.open(pth).size
                max_w, max_h = 4.6, 4.8
                w_in = min(max_w, (iw / ih) * max_h)
                h_in = w_in * ih / iw
                pic_x = Inches(12.75 - w_in)
                pic_y = Inches(3.7 - h_in / 2 + 0.4)
                pic = slide.shapes.add_picture(pth, pic_x, pic_y, width=Inches(w_in))
                anim_specs.append({"spid": pic.shape_id, "trig": "auto", "effect": img_effect, "dur": 550, "delay": 200})

    set_transition(slide, TRANSITIONS[(i - 1) % len(TRANSITIONS)])
    build_advanced_anim(slide, anim_specs)

prs.core_properties.title = "Анализ технологий охлаждения компьютерных систем"
prs.core_properties.author = "Теддер Даниил"

# Сохраняем в оба пути
prs.save(OUT)
prs.save(OUT2)
print("Успешно сохранено:")
print(" 1)", OUT)
print(" 2)", OUT2)
print("Слайдов:", TOTAL)

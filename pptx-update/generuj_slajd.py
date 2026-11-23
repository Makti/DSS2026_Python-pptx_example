"""Draft slajdu: wejścia do sieci siłowni (dane przykładowe) -- python-pptx."""
import csv, sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_MARKER_STYLE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

WEJSCIA_CSV = sys.argv[1] if len(sys.argv) > 1 else "silownia_bez_basenu_wejscia_miesiecznie.csv"
STREFY_CSV  = sys.argv[2] if len(sys.argv) > 2 else "silownia_bez_basenu_strefy_2026-10.csv"
WYNIK       = sys.argv[3] if len(sys.argv) > 3 else "slajd_silownia.pptx"

CIEMNY = RGBColor(0x26, 0x32, 0x38)
SZARY  = RGBColor(0x9E, 0xA7, 0xAD)
JASNY  = RGBColor(0xE3, 0xE7, 0xEA)
AKCENT = RGBColor(0xE4, 0x57, 0x2E)   # pomarańcz "energii"
STREFY_KOLORY = [AKCENT, CIEMNY, RGBColor(0xF2, 0xA6, 0x5A),
                 RGBColor(0x3A, 0x8D, 0xB5),   # basen
                 SZARY]                        # sauna / kolejne
FONT = "Calibri"

PELNE = dict(sty="styczeń", lut="luty", mar="marzec", kwi="kwiecień", maj="maj", cze="czerwiec",
             lip="lipiec", sie="sierpień", wrz="wrzesień", paź="październik", lis="listopad", gru="grudzień")

def fmt(x, dec=1):
    return f"{x:.{dec}f}".replace(".", ",")

# --- dane ---
with open(WEJSCIA_CSV, encoding="utf-8") as f:
    wiersze = list(csv.DictReader(f))
miesiace = [w["miesiac"] for w in wiersze]
r2025 = [float(w["wejscia_tys_2025"]) for w in wiersze]
r2026 = [float(w["wejscia_tys_2026"]) if w["wejscia_tys_2026"] else None for w in wiersze]
ost = max(i for i, v in enumerate(r2026) if v is not None)       # ostatni miesiąc z danymi
rr = (r2026[ost] / r2025[ost] - 1) * 100

with open(STREFY_CSV, encoding="utf-8") as f:
    strefy = [(w["strefa"], float(w["udzial_proc"])) for w in csv.DictReader(f)]

# --- slajd ---
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])  # pusty

def tekst(name, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    for i, linia in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for txt, size, color, bold in linia:
            r = p.add_run()
            r.text = txt
            r.font.name, r.font.size, r.font.color.rgb, r.font.bold = FONT, Pt(size), color, bold
    return tb

# Tytuł = wniosek
tekst("Tytul", 0.6, 0.45, 12.1, 0.8,
      [[(f"{PELNE[miesiace[ost]].capitalize()}: {fmt(r2026[ost])} tys. wejść, ", 34, CIEMNY, True),
        (f"{'+' if rr >= 0 else ''}{fmt(rr)}% r/r", 34, AKCENT, True)]])
tekst("Podtytul", 0.6, 1.25, 12.1, 0.4,
      [[(f"Wolne ciężary to już {strefy[0][1]:.0f}% ruchu w klubach", 18, SZARY, False)]])

# --- wykres liniowy ---
cd = CategoryChartData()
cd.categories = miesiace
cd.add_series("2025", r2025)
cd.add_series("2026", r2026)
gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(0.45), Inches(2.0),
                            Inches(7.9), Inches(4.7), cd)
gf.name = "Wykres_wejscia"
ch = gf.chart
ch.has_title = True
ch.chart_title.text_frame.text = "Wejścia miesięcznie (tys.)"
tp = ch.chart_title.text_frame.paragraphs[0]
tp.font.size, tp.font.bold, tp.font.color.rgb, tp.font.name = Pt(14), False, SZARY, FONT
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.TOP
ch.legend.include_in_layout = False
ch.legend.font.size, ch.legend.font.color.rgb, ch.legend.font.name = Pt(12), CIEMNY, FONT
for s, kolor, grub in zip(ch.plots[0].series, [SZARY, AKCENT], [2.25, 3.5]):
    s.smooth = False
    s.format.line.color.rgb = kolor
    s.format.line.width = Pt(grub)
    s.marker.style = XL_MARKER_STYLE.NONE
pt = ch.plots[0].series[1].points[ost]
pt.marker.style = XL_MARKER_STYLE.CIRCLE
pt.marker.size = 9
pt.marker.format.fill.solid(); pt.marker.format.fill.fore_color.rgb = AKCENT
pt.marker.format.line.color.rgb = AKCENT
dl = pt.data_label
dl.has_text_frame = True
dl.text_frame.text = fmt(r2026[ost])
dl.position = XL_LABEL_POSITION.ABOVE
f0 = dl.text_frame.paragraphs[0].font
f0.size, f0.bold, f0.color.rgb, f0.name = Pt(13), True, AKCENT, FONT
va, ca = ch.value_axis, ch.category_axis
va.has_major_gridlines = True
va.major_gridlines.format.line.color.rgb = JASNY
va.format.line.fill.background()
va.minimum_scale = 0
for ax in (va, ca):
    ax.tick_labels.font.size, ax.tick_labels.font.color.rgb, ax.tick_labels.font.name = Pt(11), SZARY, FONT
ca.format.line.color.rgb = JASNY

# --- wykres pierścieniowy ---
cd2 = CategoryChartData()
cd2.categories = [n for n, _ in strefy]
cd2.add_series("Udział", [v for _, v in strefy])
# przy większej liczbie stref mniejszy pierścień i gęstsza legenda, żeby zmieścić się nad źródłem
x0, y0 = 8.85, 2.0
rozm, krok = (3.7, 0.32) if len(strefy) <= 3 else (3.3, 0.28)
gf2 = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(x0), Inches(y0),
                             Inches(rozm), Inches(rozm), cd2)
gf2.name = "Wykres_strefy"
ch2 = gf2.chart
ch2.has_legend = False
ch2.has_title = False
pl = ch2.plots[0]
pl.vary_by_categories = True
hole = pl._element.find("{http://schemas.openxmlformats.org/drawingml/2006/chart}holeSize")
if hole is not None:
    hole.set("val", "62")
for i, p in enumerate(pl.series[0].points):
    p.format.fill.solid(); p.format.fill.fore_color.rgb = STREFY_KOLORY[i % len(STREFY_KOLORY)]
    p.format.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); p.format.line.width = Pt(2)

# KPI w środku pierścienia
tekst("KPI_srodek", x0 + 0.6, y0 + rozm / 2 - 0.55, rozm - 1.2, 1.1,
      [[(f"{strefy[0][1]:.0f}%", 40, AKCENT, True)],
       [(strefy[0][0].lower(), 13, SZARY, False)]],
      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Legenda pod pierścieniem (ręczna -> pełna kontrola)
ly = y0 + rozm + 0.15
for i, (n, v) in enumerate(strefy):
    kw = slide.shapes.add_shape(1, Inches(x0 + 0.45), Inches(ly + i * krok + 0.06),
                                Inches(0.16), Inches(0.16))
    kw.name = f"Legenda_kolor_{i+1}"
    kw.fill.solid(); kw.fill.fore_color.rgb = STREFY_KOLORY[i % len(STREFY_KOLORY)]; kw.line.fill.background()
    tekst(f"Legenda_{i+1}", x0 + 0.75, ly + i * krok, 2.9, 0.3,
          [[(n, 13, CIEMNY, False), (f"  {v:.0f}%", 13, CIEMNY, True)]])

tekst("Zrodlo", 0.6, 7.0, 12.1, 0.3,
      [[(f"Dane przykładowe (fikcyjne). Struktura wejść: {PELNE[miesiace[ost]]} 2026.", 10, SZARY, False)]])

prs.save(WYNIK)
print("Zapisano", WYNIK)

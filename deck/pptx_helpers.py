"""Drawing helpers on top of python-pptx (all coordinates in inches)."""
import copy
import re

from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION, XL_TICK_MARK
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

NAVY = RGBColor(0x00, 0x50, 0x96)
BLUE = RGBColor(0x00, 0x78, 0xB4)
CYAN = RGBColor(0x1E, 0xB9, 0xE1)
DARK = RGBColor(0x3C, 0x3C, 0x3C)
GRAY = RGBColor(0x96, 0x96, 0x96)
LIGHT = RGBColor(0xD7, 0xD7, 0xD7)
PALE = RGBColor(0xE9, 0xF3, 0xFA)
PALE2 = RGBColor(0xF4, 0xF6, 0xF8)
PALECYAN = RGBColor(0xD8, 0xF2, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
YELLOW = RGBColor(0xFF, 0xC8, 0x19)
RED = RGBColor(0xC0, 0x39, 0x2B)

ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCHOR = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}

_BOLD = re.compile(r"(\*\*.+?\*\*)")


def add_runs(paragraph, text, size, color=DARK, bold=False, bold_color=None):
    """Write `text` into a paragraph; **x** marks bold spans."""
    for part in _BOLD.split(text):
        if not part:
            continue
        is_bold = part.startswith("**") and part.endswith("**")
        run = paragraph.add_run()
        run.text = part[2:-2] if is_bold else part
        run.font.size = Pt(size)
        run.font.bold = bool(is_bold or bold)
        run.font.color.rgb = (bold_color or color) if is_bold else color


def _fill_text(tf, text, size, color, bold, align, bold_color=None, line_spacing=None, space_after=0):
    tf.word_wrap = True
    lines = text if isinstance(text, (list, tuple)) else str(text).split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN[align]
        if line_spacing:
            p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        p.space_before = Pt(0)
        if isinstance(line, tuple):
            ltext = line[0]
            lsize = line[1] if len(line) > 1 else size
            lcolor = line[2] if len(line) > 2 else color
            lbold = line[3] if len(line) > 3 else bold
            add_runs(p, ltext, lsize, lcolor, lbold, bold_color)
        else:
            add_runs(p, line, size, color, bold, bold_color)


def _margins(tf, m):
    tf.margin_left = tf.margin_right = Inches(m)
    tf.margin_top = tf.margin_bottom = Inches(min(m, 0.03))


def tb(slide, x, y, w, h, text, size=9, color=DARK, bold=False, align="l", anchor="t",
       bold_color=None, line_spacing=None, space_after=0, margin=0.03):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    _margins(tf, margin)
    tf.vertical_anchor = ANCHOR[anchor]
    _fill_text(tf, text, size, color, bold, align, bold_color, line_spacing, space_after)
    return shape


def box(slide, x, y, w, h, text="", fill=PALE, line=None, size=9, color=DARK, bold=False,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE, align="c", anchor="m", radius=0.12, line_w=0.75,
        dash=False, bold_color=None, margin=0.05, line_spacing=None, space_after=0):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    # drop the theme style reference so no theme shadow/gradient is applied
    style = shp._element.find(qn("p:style"))
    if style is not None:
        shp._element.remove(style)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE and len(shp.adjustments):
        shp.adjustments[0] = radius
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
        if dash:
            shp.line.dash_style = MSO_LINE.DASH
    tf = shp.text_frame
    _margins(tf, margin)
    tf.vertical_anchor = ANCHOR[anchor]
    if text != "":
        _fill_text(tf, text, size, color, bold, align, bold_color, line_spacing, space_after)
    return shp


def line(slide, x1, y1, x2, y2, color=GRAY, width=1.0, arrow=False, dash=False):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(width)
    if dash:
        conn.line.dash_style = MSO_LINE.DASH
    if arrow:
        ln = conn.line._get_or_add_ln()
        tail = etree.SubElement(ln, qn("a:tailEnd"))
        tail.set("type", "triangle")
        tail.set("w", "med")
        tail.set("len", "med")
    return conn


def arrow(slide, x1, y1, x2, y2, color=GRAY, width=1.25, dash=False):
    return line(slide, x1, y1, x2, y2, color, width, arrow=True, dash=dash)


def ctitle(slide, x, y, w, text, size=9.5):
    """Section/chart title with a small navy marker."""
    box(slide, x, y + 0.03, 0.05, 0.17, fill=NAVY, shape=MSO_SHAPE.RECTANGLE)
    return tb(slide, x + 0.09, y, w - 0.09, 0.23, text, size=size, color=NAVY, bold=True, anchor="m")


def kpi(slide, x, y, w, h, big, small, big_size=15, small_size=7.5, fill=PALE):
    box(slide, x, y, w, h, fill=fill, radius=0.08)
    box(slide, x, y, 0.045, h, fill=NAVY, shape=MSO_SHAPE.RECTANGLE)
    split = h * 0.50
    tb(slide, x + 0.1, y + 0.02, w - 0.14, split, big, size=big_size, color=NAVY, bold=True, anchor="m")
    tb(slide, x + 0.1, y + split, w - 0.14, h - split - 0.02, small, size=small_size, color=DARK, anchor="t")


def source(slide, text):
    return tb(slide, 0.38, 4.93, 9.24, 0.2, "数据来源：" + text, size=7, color=GRAY, anchor="m")


def picture(slide, path, x, y, w=None, h=None):
    kwargs = {}
    if w:
        kwargs["width"] = Inches(w)
    if h:
        kwargs["height"] = Inches(h)
    return slide.shapes.add_picture(path, Inches(x), Inches(y), **kwargs)


def bullets(ph, items, x, y, w, h, size=9, space_after=3, bold_color=NAVY, line_spacing=1.08):
    """Fill the layout's content placeholder with bulleted text. Items: str or (level, str)."""
    ph.left, ph.top, ph.width, ph.height = Inches(x), Inches(y), Inches(w), Inches(h)
    tf = ph.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, item in enumerate(items):
        level, text = item if isinstance(item, tuple) else (0, item)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        ppr = p._p.get_or_add_pPr()
        indent = 171450
        ppr.set("marL", str(indent * (level + 1)))
        ppr.set("indent", str(-indent))
        p.space_after = Pt(space_after)
        p.space_before = Pt(0)
        p.line_spacing = line_spacing
        bu_font = etree.SubElement(ppr, qn("a:buFont"))
        bu_font.set("typeface", "Arial")
        bu_char = etree.SubElement(ppr, qn("a:buChar"))
        bu_char.set("char", "•" if level == 0 else "–")
        add_runs(p, text, size - (0.5 if level else 0), DARK, False, bold_color)
    return ph


def remove_shape(shape):
    shape._element.getparent().remove(shape._element)


# ---------------------------------------------------------------- charts

def _style_chart(chart, fs, legend):
    chart.has_title = False
    chart.font.size = Pt(fs)
    chart.font.color.rgb = DARK
    chart.has_legend = bool(legend)
    if legend:
        chart.legend.position = {"b": XL_LEGEND_POSITION.BOTTOM, "t": XL_LEGEND_POSITION.TOP,
                                 "r": XL_LEGEND_POSITION.RIGHT}[legend]
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(fs)


def _style_axes(chart, fs, vmin=None, vmax=None, labels_low=False):
    va = chart.value_axis
    va.visible = False
    va.has_major_gridlines = False
    va.has_minor_gridlines = False
    if vmin is not None:
        va.minimum_scale = vmin
    if vmax is not None:
        va.maximum_scale = vmax
    ca = chart.category_axis
    ca.has_major_gridlines = False
    ca.major_tick_mark = XL_TICK_MARK.NONE
    ca.format.line.color.rgb = LIGHT
    ca.tick_labels.font.size = Pt(fs)
    if labels_low:
        ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW


def _custom_labels(series, texts, fs, color=DARK):
    for i, text in enumerate(texts):
        if text is None:
            continue
        tf = series.points[i].data_label.text_frame
        tf.text = text
        for p in tf.paragraphs:
            for r in p.runs:
                r.font.size = Pt(fs)
                r.font.color.rgb = color


def bar_chart(slide, x, y, w, h, cats, series, colors, horizontal=False, stacked=False, fmt="#,##0",
              fs=8, legend=None, gap=55, overlap=None, point_colors=None, vmin=None, vmax=None,
              labels=True, custom_labels=None, labels_low=False, label_fs=None):
    """series: list of (name, values). custom_labels: {series_idx: [text...]}"""
    data = CategoryChartData()
    data.categories = cats
    for name, values in series:
        data.add_series(name, values)
    if horizontal:
        ctype = XL_CHART_TYPE.BAR_STACKED if stacked else XL_CHART_TYPE.BAR_CLUSTERED
    else:
        ctype = XL_CHART_TYPE.COLUMN_STACKED if stacked else XL_CHART_TYPE.COLUMN_CLUSTERED
    gf = slide.shapes.add_chart(ctype, Inches(x), Inches(y), Inches(w), Inches(h), data)
    chart = gf.chart
    _style_chart(chart, fs, legend)
    _style_axes(chart, fs, vmin, vmax, labels_low)
    if horizontal:
        chart.category_axis.reverse_order = True
    plot = chart.plots[0]
    plot.gap_width = gap
    if overlap is not None:
        plot.overlap = overlap
    plot.has_data_labels = bool(labels)
    if labels:
        dl = plot.data_labels
        dl.number_format = fmt
        dl.number_format_is_linked = False
        dl.font.size = Pt(label_fs or fs)
        dl.font.color.rgb = DARK
        if not stacked:
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
    for i, s in enumerate(plot.series):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = colors[i % len(colors)]
        s.invert_if_negative = False
        if point_colors and i == 0:
            for j, c in enumerate(point_colors):
                pt = s.points[j]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = c
        if custom_labels and i in custom_labels:
            _custom_labels(s, custom_labels[i], label_fs or fs)
    return chart


def line_chart(slide, x, y, w, h, cats, series, colors, fmt="0.0", fs=8, legend=None, vmin=None, vmax=None,
               labels=True, custom_labels=None, labels_low=False, label_pos=XL_LABEL_POSITION.ABOVE, width=2.0):
    data = CategoryChartData()
    data.categories = cats
    for name, values in series:
        data.add_series(name, values)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(x), Inches(y), Inches(w), Inches(h), data)
    chart = gf.chart
    _style_chart(chart, fs, legend)
    _style_axes(chart, fs, vmin, vmax, labels_low)
    plot = chart.plots[0]
    plot.has_data_labels = bool(labels)
    if labels:
        dl = plot.data_labels
        dl.number_format = fmt
        dl.number_format_is_linked = False
        dl.font.size = Pt(fs)
        dl.font.color.rgb = DARK
        dl.position = label_pos
    for i, s in enumerate(plot.series):
        c = colors[i % len(colors)]
        s.smooth = False
        s.format.line.color.rgb = c
        s.format.line.width = Pt(width)
        s.marker.style = 8  # circle
        s.marker.size = 6
        s.marker.format.fill.solid()
        s.marker.format.fill.fore_color.rgb = c
        s.marker.format.line.color.rgb = c
        if custom_labels and i in custom_labels:
            _custom_labels(s, custom_labels[i], fs)
    return chart


def doughnut(slide, x, y, w, h, cats, values, colors, fs=8, hole=58, legend="r"):
    data = CategoryChartData()
    data.categories = cats
    data.add_series("share", values)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(x), Inches(y), Inches(w), Inches(h), data)
    chart = gf.chart
    _style_chart(chart, fs, legend)
    plot = chart.plots[0]
    plot.has_data_labels = False
    s = plot.series[0]
    for j, c in enumerate(colors):
        pt = s.points[j]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = c
        pt.format.line.color.rgb = WHITE
        pt.format.line.width = Pt(0.75)
    hole_el = plot._element.find(qn("c:holeSize"))
    if hole_el is None:
        hole_el = etree.SubElement(plot._element, qn("c:holeSize"))
    hole_el.set("val", str(hole))
    return chart


# ---------------------------------------------------------------- tables

_NO_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"


def _cell_border(cell, color="D7D7D7", width_pt=0.5, sides=("B",)):
    tc_pr = cell._tc.get_or_add_tcPr()
    for tag in ("lnL", "lnR", "lnT", "lnB"):
        old = tc_pr.find(qn("a:" + tag))
        if old is not None:
            tc_pr.remove(old)
    pos = 0
    for side in ("L", "R", "T", "B"):
        ln = etree.Element(qn("a:ln" + side))
        if side in sides:
            ln.set("w", str(int(width_pt * 12700)))
            ln.set("cap", "flat")
            ln.set("cmpd", "sng")
            fill = etree.SubElement(ln, qn("a:solidFill"))
            clr = etree.SubElement(fill, qn("a:srgbClr"))
            clr.set("val", color)
        else:
            ln.set("w", "0")
            etree.SubElement(ln, qn("a:noFill"))
        tc_pr.insert(pos, ln)
        pos += 1


def table(slide, x, y, colw, rows, fs=8, rowh=0.25, header_fill=NAVY, aligns=None, zebra=True,
          bold_first_col=False, highlight_rows=(), header_fs=None, bold_color=NAVY, first_col_fill=None):
    """rows[0] is the header. Cell text may contain \\n and **bold**."""
    nrows, ncols = len(rows), len(colw)
    gf = slide.shapes.add_table(nrows, ncols, Inches(x), Inches(y), Inches(sum(colw)), Inches(rowh * nrows))
    tbl = gf.table
    style = tbl._tbl.tblPr.find(qn("a:tableStyleId"))
    if style is not None:
        style.text = _NO_STYLE
    tbl.first_row = True
    tbl.horz_banding = False
    for j, cw in enumerate(colw):
        tbl.columns[j].width = Inches(cw)
    aligns = aligns or ["l"] + ["c"] * (ncols - 1)
    for i, row in enumerate(rows):
        tbl.rows[i].height = Inches(rowh)
        for j in range(ncols):
            cell = tbl.cell(i, j)
            text = row[j] if j < len(row) else ""
            cell.margin_left = cell.margin_right = Inches(0.05)
            cell.margin_top = cell.margin_bottom = Inches(0.01)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if i == 0:
                cell.fill.fore_color.rgb = header_fill
                color, bold, size = WHITE, True, header_fs or fs
            else:
                if i in highlight_rows:
                    cell.fill.fore_color.rgb = PALE
                elif j == 0 and first_col_fill is not None:
                    cell.fill.fore_color.rgb = first_col_fill
                elif zebra and i % 2 == 0:
                    cell.fill.fore_color.rgb = PALE2
                else:
                    cell.fill.fore_color.rgb = WHITE
                color, size = DARK, fs
                bold = (bold_first_col and j == 0) or (i in highlight_rows)
            tf = cell.text_frame
            tf.word_wrap = True
            lines = str(text).split("\n")
            for k, ln_text in enumerate(lines):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.alignment = ALIGN[aligns[j] if i > 0 else ("l" if j == 0 and aligns[0] == "l" else "c")]
                add_runs(p, ln_text, size, color, bold, WHITE if i == 0 else bold_color)
            _cell_border(cell, "D7D7D7", 0.5, ("B",))
    return gf


def move_slide(prs, old_index, new_index):
    sld_id_lst = prs.slides._sldIdLst
    items = list(sld_id_lst)
    el = items[old_index]
    sld_id_lst.remove(el)
    sld_id_lst.insert(new_index, el)


def clone_placeholder_from(slide_src, slide_dst, idx=10):
    for shp in slide_src.placeholders:
        if shp.placeholder_format.idx == idx:
            el = copy.deepcopy(shp._element)
            slide_dst.shapes._spTree.append(el)
            return slide_dst.shapes[-1]
    return None

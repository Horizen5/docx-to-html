# -*- coding: utf-8 -*-
"""
DOCX → HTML 通用转换器
接受任意 .docx 文件，自动解压并 1:1 还原为 HTML

用法:
  python docx2html.py input.docx                    # 输出到同目录
  python docx2html.py input.docx -o output.html     # 指定输出路径
  python docx2html.py --web                          # 启动 Web 上传界面
"""
import xml.etree.ElementTree as ET
import base64
import os
import sys
import html as html_lib
import math
import zipfile
import tempfile
import shutil
import argparse

# ===== 命名空间 =====
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wp14':'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
}

W   = NS['w']
A   = NS['a']
WPS = NS['wps']
WPG = NS['wpg']
WP  = NS['wp']

EMU_PER_MM = 36000
EMU_PER_PT = 12700


# ===================================================================
#  工具函数
# ===================================================================
def gattr(elem, attr, ns_prefix=None):
    """安全获取属性，先尝试带命名空间"""
    if elem is None:
        return None
    if ns_prefix:
        val = elem.get(f'{{{NS[ns_prefix]}}}{attr}')
        if val is not None:
            return val
    return elem.get(attr)

def gchild(parent, name, ns_prefix):
    if parent is None:
        return None
    return parent.find(f'{{{NS[ns_prefix]}}}{name}')

def gchildren(parent, name, ns_prefix):
    if parent is None:
        return []
    return parent.findall(f'{{{NS[ns_prefix]}}}{name}')

def gtext(elem):
    if elem is None:
        return ''
    return elem.text or ''

def emu_to_mm(emu):
    return round(int(emu) / EMU_PER_MM, 3)

def parse_color(color_val):
    if not color_val or color_val == 'auto':
        return None
    return f'#{color_val.upper()}'

def hex_to_rgba(hex_color, alpha=1):
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return f'rgba({r},{g},{b},{alpha:.2f})'


# ===================================================================
#  主题颜色加载
# ===================================================================
def load_theme_colors(docx_dir):
    """从 theme1.xml 读取主题颜色方案"""
    theme_path = os.path.join(docx_dir, 'word', 'theme', 'theme1.xml')
    colors = {
        'tx1': '#000000', 'tx2': '#1F497D',
        'bg1': '#FFFFFF', 'bg2': '#EEECE1',
        'accent1': '#4F81BD', 'accent2': '#C0504D',
        'accent3': '#9BBB59', 'accent4': '#8064A2',
        'accent5': '#4BACC6', 'accent6': '#F79646',
        'dk1': '#000000', 'lt1': '#FFFFFF',
        'dk2': '#1F497D', 'lt2': '#EEECE1',
        'hlink': '#0000FF', 'folHlink': '#800080',
    }
    if not os.path.exists(theme_path):
        return colors

    try:
        tree = ET.parse(theme_path)
        root = tree.getroot()
        clr_scheme = root.find(f'.//{{{A}}}clrScheme')
        if clr_scheme is None:
            return colors

        for child in clr_scheme:
            tag = child.tag.split('}')[-1]
            # sysClr (dk1, lt1)
            sys_clr = child.find(f'{{{A}}}sysClr')
            if sys_clr is not None:
                last_clr = gattr(sys_clr, 'lastClr', 'a')
                if last_clr:
                    colors[tag] = f'#{last_clr.upper()}'
                continue
            # srgbClr (dk2, lt2, accent1-6, hlink, folHlink)
            srgb = child.find(f'{{{A}}}srgbClr')
            if srgb is not None:
                val = gattr(srgb, 'val', 'a')
                if val:
                    colors[tag] = f'#{val.upper()}'
    except Exception:
        pass

    return colors

THEME_COLORS = {}


# ---- 颜色变换辅助（处理 schemeClr 的 lumMod/lumOff/tint/shade 等子元素）----
def _hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)

def _rgb_to_hex(r, g, b):
    def _c(v):
        return max(0, min(255, int(round(v))))
    return f'#{_c(r):02X}{_c(g):02X}{_c(b):02X}'

def _rgb_to_hsl(r, g, b):
    r, g, b = r / 255, g / 255, b / 255
    mx, mn = max(r, g, b), min(r, g, b)
    l = (mx + mn) / 2
    if mx == mn:
        h = s = 0.0
    else:
        d = mx - mn
        s = d / (2 - mx - mn) if l > 0.5 else d / (mx + mn)
        if mx == r:
            h = (g - b) / d + (6 if g < b else 0)
        elif mx == g:
            h = (b - r) / d + 2
        else:
            h = (r - g) / d + 4
        h /= 6
    return h, s, l

def _hsl_to_rgb(h, s, l):
    if s == 0:
        r = g = b = l
    else:
        def hue2rgb(p, q, t):
            if t < 0:
                t += 1
            if t > 1:
                t -= 1
            if t < 1 / 6:
                return p + (q - p) * 6 * t
            if t < 1 / 2:
                return q
            if t < 2 / 3:
                return p + (q - p) * (2 / 3 - t) * 6
            return p
        q = l + s * (1 - l) if l < 0.5 else l + s - l * s
        p = 2 * l - q
        r = hue2rgb(p, q, h + 1 / 3)
        g = hue2rgb(p, q, h)
        b = hue2rgb(p, q, h - 1 / 3)
    return r * 255, g * 255, b * 255

def _adjust_luminance(hex_color, factor, mode):
    """mode: 'mod' 乘算亮度, 'off' 加算亮度。factor 为 0~1"""
    r, g, b = _hex_to_rgb(hex_color)
    h, s, l = _rgb_to_hsl(r, g, b)
    if mode == 'mod':
        l = l * factor
    elif mode == 'off':
        l = l + factor
    l = max(0.0, min(1.0, l))
    r, g, b = _hsl_to_rgb(h, s, l)
    return _rgb_to_hex(r, g, b)

def _blend(hex_color, target, t):
    r, g, b = _hex_to_rgb(hex_color)
    tr, tg, tb = target
    return _rgb_to_hex(r + (tr - r) * t, g + (tg - g) * t, b + (tb - b) * t)

def parse_scheme_clr(scheme_elem):
    """解析 <a:schemeClr val="tx1"> 元素，含 lumMod/lumOff/tint/shade 等子元素修饰。"""
    if scheme_elem is None:
        return None
    val = gattr(scheme_elem, 'val', 'a')
    base = THEME_COLORS.get(val, '#000000')
    for mod in scheme_elem:
        local = mod.tag.split('}')[-1] if '}' in mod.tag else mod.tag
        mval = gattr(mod, 'val', 'a')
        if mval is None:
            continue
        factor = int(mval) / 100000.0
        if local == 'lumMod':
            base = _adjust_luminance(base, factor, 'mod')
        elif local == 'lumOff':
            base = _adjust_luminance(base, factor, 'off')
        elif local == 'tint':
            base = _blend(base, (255, 255, 255), factor)
        elif local == 'shade':
            base = _blend(base, (0, 0, 0), factor)
    return base

def parse_scheme(scheme):
    return THEME_COLORS.get(scheme, '#000000')


# ===================================================================
#  文本运行 / 段落解析
# ===================================================================
def parse_run(run_elem):
    if run_elem is None:
        return '', {}

    rpr = gchild(run_elem, 'rPr', 'w')
    style = {}

    if rpr is not None:
        rfonts = gchild(rpr, 'rFonts', 'w')
        if rfonts is not None:
            ascii_font = gattr(rfonts, 'ascii', 'w')
            ea_font = gattr(rfonts, 'eastAsia', 'w')
            if ascii_font:
                style['font-family'] = f"'{ascii_font}', sans-serif"
            elif ea_font:
                style['font-family'] = f"'{ea_font}', sans-serif"

        color_elem = gchild(rpr, 'color', 'w')
        if color_elem is not None:
            c = parse_color(gattr(color_elem, 'val', 'w'))
            if c:
                style['color'] = c

        sz_elem = gchild(rpr, 'sz', 'w')
        if sz_elem is not None:
            sz_val = int(gattr(sz_elem, 'val', 'w') or '24') / 2.0
            style['font-size'] = f'{sz_val}pt'

        b_elem = gchild(rpr, 'b', 'w')
        if b_elem is not None:
            b_val = gattr(b_elem, 'val', 'w')
            style['font-weight'] = 'bold' if b_val != '0' else 'normal'

        i_elem = gchild(rpr, 'i', 'w')
        if i_elem is not None:
            i_val = gattr(i_elem, 'val', 'w')
            style['font-style'] = 'italic' if i_val != '0' else 'normal'

        u_elem = gchild(rpr, 'u', 'w')
        if u_elem is not None:
            u_val = gattr(u_elem, 'val', 'w')
            if u_val and u_val != 'none':
                style['text-decoration'] = 'underline'

        # w14 阴影
        shadow = gchild(rpr, 'shadow', 'w14')
        if shadow is not None:
            blur = int(gattr(shadow, 'blurRad', 'w14') or '0') / EMU_PER_PT
            dist = int(gattr(shadow, 'dist', 'w14') or '0') / EMU_PER_PT
            dir_val = int(gattr(shadow, 'dir', 'w14') or '0')
            srgb = gchild(shadow, 'srgbClr', 'w14')
            if srgb is not None:
                alpha_elem = gchild(srgb, 'alpha', 'w14')
                alpha_val = int(gattr(alpha_elem, 'val', 'w14') or '100000') / 100000 if alpha_elem is not None else 1
                shadow_color = '#' + (gattr(srgb, 'val', 'w14') or '000000').upper()
            else:
                shadow_color = '#000000'
                alpha_val = 1
            rad = math.radians(dir_val / 60000)
            dx = dist * math.cos(rad)
            dy = dist * math.sin(rad)
            rgba = hex_to_rgba(shadow_color, alpha_val)
            style['text-shadow'] = f'{dx:.1f}pt {dy:.1f}pt {blur:.1f}pt {rgba}'

    # 收集文本（仅直接子元素，不递归到 drawing/textbox 内部）
    text_parts = []
    for child in run_elem:
        local = child.tag.split('}')[-1]
        if local == 't':
            text_parts.append(gtext(child))
        elif local == 'tab':
            text_parts.append('\t')
        elif local == 'br':
            br_type = gattr(child, 'type', 'w')
            if br_type == 'page':
                text_parts.append('\f')  # form feed = page break
            else:
                text_parts.append('<br>')

    return ''.join(text_parts), style


def parse_spacing(spacing_elem):
    """解析 w:spacing 元素，返回 (line_height_str, before_pt, after_pt)"""
    if spacing_elem is None:
        return None, 0, 0

    sl = gattr(spacing_elem, 'line', 'w')
    sl_rule = gattr(spacing_elem, 'lineRule', 'w') or 'auto'
    spacing_line = None

    if sl:
        sl_val = int(sl)
        if sl_rule == 'exact':
            spacing_line = f'{sl_val / 20.0}pt'
        elif sl_rule == 'atLeast':
            if sl_val / 20.0 >= 10:
                spacing_line = f'{sl_val / 20.0}pt'
        else:
            multiplier = sl_val / 240.0
            if multiplier > 0 and abs(multiplier - 1.0) > 0.01:
                spacing_line = f'{multiplier:.2f}'

    sb = gattr(spacing_elem, 'before', 'w')
    spacing_before = int(sb) / 20.0 if sb else 0
    sa = gattr(spacing_elem, 'after', 'w')
    spacing_after = int(sa) / 20.0 if sa else 0

    return spacing_line, spacing_before, spacing_after


def parse_tab_stops(ppr):
    """解析段落的 <w:tabs> 制表位，返回相对段落左边的位置列表（单位 mm）。"""
    if ppr is None:
        return []
    tabs = gchild(ppr, 'tabs', 'w')
    if tabs is None:
        return []
    stops = []
    for tab in gchildren(tabs, 'tab', 'w'):
        pos = gattr(tab, 'pos', 'w')
        if pos:
            # twips -> mm
            stops.append(int(pos) / 1440.0 * 25.4)
    return stops


def _render_tabbed(text, style_str, tab_stops, tab_state):
    """将一段文本中的 \\t 按真实制表位宽度渲染为占位 span。"""
    if '\t' not in text:
        escaped = html_lib.escape(text)
        if style_str:
            return f'<span style="{style_str}">{escaped}</span>'
        return f'<span>{escaped}</span>'

    parts = text.split('\t')
    out = []
    for i, part in enumerate(parts):
        if part:
            escaped = html_lib.escape(part)
            if style_str:
                out.append(f'<span style="{style_str}">{escaped}</span>')
            else:
                out.append(f'<span>{escaped}</span>')
        if i < len(parts) - 1:
            # 在制表位之间插入相当于列间距的占位
            if tab_state['idx'] < len(tab_stops):
                gap = tab_stops[tab_state['idx']] if tab_state['idx'] == 0 \
                    else tab_stops[tab_state['idx']] - tab_stops[tab_state['idx'] - 1]
                tab_state['idx'] += 1
                out.append(f'<span style="display:inline-block;width:{gap:.1f}mm;"></span>')
            else:
                out.append('&nbsp;&nbsp;&nbsp;&nbsp;')
    return ''.join(out)


# Word 的正文段落常常不写任何 <w:spacing>，但 Word 应用本身仍会对 Normal 样式
# 施加一个默认的「段后距」（space-after）。浏览器默认 margin:0 会把这个高度丢干净，
# 导致整页被纵向压缩、章节条/照片整体向上漂移几十毫米。
# 这里主动补一个与 Word 默认一致的段后距，消除漂移。取值经与 PDF 真值逐带比对校准。
DEFAULT_SPACE_AFTER_PT = 5.0
# 带背景形状的「小节条/章节标题」（如简历里的灰色章节条）通常靠下方空段落撑出间隙，
# 但浏览器里空段落会塌缩为 0 高度，导致条紧贴正文。这里给这类有显式精确行高的
# 段落补一个段后距，使条与下方正文之间留出与 PDF 一致间隙（值经 PDF 真值比对校准）。
# 注意：小节条下方加段后距会向下级联，故同步下调 DEFAULT_SPACE_AFTER_PT 抵消，
# 整体纵向位置仍对齐 PDF 真值（Word 正文段后距本就约 3~4pt，而非 6.5pt）。
BAR_SPACE_AFTER_PT = 8.0


def parse_paragraph(p_elem):
    """解析段落，返回 HTML 字符串。可能包含分页符 \\f"""
    if p_elem is None:
        return ''

    ppr = gchild(p_elem, 'pPr', 'w')
    align = None
    spacing_line = None
    spacing_before = 0
    spacing_after = 0
    tab_stops = []
    has_spacing_elem = False

    if ppr is not None:
        jc = gchild(ppr, 'jc', 'w')
        if jc is not None:
            align = gattr(jc, 'val', 'w')
        spacing = gchild(ppr, 'spacing', 'w')
        has_spacing_elem = spacing is not None
        spacing_line, spacing_before, spacing_after = parse_spacing(spacing)
        tab_stops = parse_tab_stops(ppr)

    # 检测段落最大字号，用于给大字号标题/横幅施加更宽松的行距（贴近 Word 标题行高）。
    max_sz = 0
    for _r in gchildren(p_elem, 'r', 'w'):
        for _sz in _r.findall(f'.//{{{W}}}sz'):
            try:
                max_sz = max(max_sz, int(gattr(_sz, 'val', 'w') or 0))
            except ValueError:
                pass

    runs = gchildren(p_elem, 'r', 'w')
    if not runs:
        # 检查是否有 hyperlink
        hyperlinks = gchildren(p_elem, 'hyperlink', 'w')
        for hl in hyperlinks:
            hl_runs = gchildren(hl, 'r', 'w')
            runs.extend(hl_runs)
    if not runs:
        return ''

    tab_state = {'idx': 0}
    run_html_parts = []
    for run in runs:
        text, style = parse_run(run)
        if not text:
            continue
        style_str = '; '.join(f'{k}: {v}' for k, v in style.items()) if style else ''

        # 处理分页符
        if '\f' in text:
            parts = text.split('\f')
            for i, part in enumerate(parts):
                run_html_parts.append(_render_tabbed(part, style_str, tab_stops, tab_state))
                if i < len(parts) - 1:
                    run_html_parts.append('<div style="page-break-after: always;"></div>')
            continue

        # 处理 <br>
        if '<br>' in text:
            parts = text.split('<br>')
            for i, part in enumerate(parts):
                run_html_parts.append(_render_tabbed(part, style_str, tab_stops, tab_state))
                if i < len(parts) - 1:
                    run_html_parts.append('<br>')
            continue

        run_html_parts.append(_render_tabbed(text, style_str, tab_stops, tab_state))

    if not run_html_parts:
        return ''

    p_style_parts = []
    if align:
        p_style_parts.append(f'text-align: {align}')
    # 行距：优先 docx 显式 w:spacing 的 line；否则正文继承 body 的 1.40×。
    # 大字号标题/横幅（≥24pt）用更宽松的 1.60×，贴近 Word 标题行高（否则横幅偏矮）。
    if spacing_line:
        p_style_parts.append(f'line-height: {spacing_line}')
    elif max_sz >= 48:
        p_style_parts.append('line-height: 1.60')
    # 上下边距：优先用 docx 显式 <w:spacing> 的 before/after；
    # 若段落完全没有任何 <w:spacing>（Word 正文常态），浏览器 margin:0 会丢失
    # Word 默认段后距 → 整页纵向压缩。这里补一个与 Word 默认一致的段后距；
    # 上边距归零，避免浏览器默认 1em 顶距。段间靠 CSS margin 折叠（与 Word 一致，
    # 相邻段只保留较大者而非相加），故对所有段统一施加也不会重复累加。
    if spacing_before:
        p_style_parts.append(f'margin-top: {spacing_before}pt')
    else:
        p_style_parts.append('margin-top: 0')
    if spacing_after:
        p_style_parts.append(f'margin-bottom: {spacing_after}pt')
    elif not has_spacing_elem:
        p_style_parts.append(f'margin-bottom: {DEFAULT_SPACE_AFTER_PT}pt')
    else:
        # 段落有显式 <w:spacing> 但无 after：多为带背景形状的「小节条/标题」
        # （如简历灰色章节条，有精确行高但下方空段落塌缩）。补一个段后距留出间隙。
        if spacing_line is not None:
            p_style_parts.append(f'margin-bottom: {BAR_SPACE_AFTER_PT}pt')
        else:
            p_style_parts.append('margin-bottom: 0')
    p_style_parts.append('padding: 0')

    p_style = '; '.join(p_style_parts)
    inner = ''.join(run_html_parts)
    return f'<p style="{p_style}">{inner}</p>'


def parse_txbx_content(txbx_elem):
    if txbx_elem is None:
        return ''
    paras = gchildren(txbx_elem, 'p', 'w')
    html_parts = []
    for p in paras:
        h = parse_paragraph(p)
        if h:
            html_parts.append(h)
    return ''.join(html_parts)


# ===================================================================
#  几何 → SVG
# ===================================================================
def parse_cust_geom_to_svg(cust_geom, fill_color, line_color, line_w_pt, no_fill, no_line):
    if cust_geom is None:
        return None

    path_lst = gchild(cust_geom, 'pathLst', 'a')
    if path_lst is None:
        return None

    svg_paths = []
    viewBox_w = 1
    viewBox_h = 1

    for path_elem in path_lst:
        if path_elem.tag != f'{{{A}}}path':
            continue
        path_w = int(gattr(path_elem, 'w', 'a') or '1')
        path_h = int(gattr(path_elem, 'h', 'a') or '1')
        viewBox_w = max(viewBox_w, path_w)
        viewBox_h = max(viewBox_h, path_h)

        d_parts = []
        for cmd in path_elem:
            local = cmd.tag.split('}')[-1] if '}' in cmd.tag else cmd.tag

            if local == 'moveTo':
                pts = gchildren(cmd, 'pt', 'a')
                if pts:
                    x = gattr(pts[0], 'x', 'a') or '0'
                    y = gattr(pts[0], 'y', 'a') or '0'
                    d_parts.append(f'M {x} {y}')
            elif local == 'lnTo':
                pts = gchildren(cmd, 'pt', 'a')
                if pts:
                    x = gattr(pts[0], 'x', 'a') or '0'
                    y = gattr(pts[0], 'y', 'a') or '0'
                    d_parts.append(f'L {x} {y}')
            elif local == 'cubicBezTo':
                pts = gchildren(cmd, 'pt', 'a')
                if len(pts) >= 3:
                    x1, y1 = gattr(pts[0], 'x', 'a') or '0', gattr(pts[0], 'y', 'a') or '0'
                    x2, y2 = gattr(pts[1], 'x', 'a') or '0', gattr(pts[1], 'y', 'a') or '0'
                    x3, y3 = gattr(pts[2], 'x', 'a') or '0', gattr(pts[2], 'y', 'a') or '0'
                    d_parts.append(f'C {x1} {y1} {x2} {y2} {x3} {y3}')
            elif local == 'quadBezTo':
                pts = gchildren(cmd, 'pt', 'a')
                if len(pts) >= 2:
                    x1, y1 = gattr(pts[0], 'x', 'a') or '0', gattr(pts[0], 'y', 'a') or '0'
                    x2, y2 = gattr(pts[1], 'x', 'a') or '0', gattr(pts[1], 'y', 'a') or '0'
                    d_parts.append(f'Q {x1} {y1} {x2} {y2}')
            elif local == 'close':
                d_parts.append('Z')

        if d_parts:
            svg_paths.append(' '.join(d_parts))

    if not svg_paths:
        return None

    fill_attr = 'none' if no_fill else (fill_color or 'none')
    stroke_attr = 'none' if no_line else (line_color or 'none')
    stroke_w = f'{line_w_pt}pt' if line_w_pt and not no_line else '0'

    path_data = ' '.join(svg_paths)

    svg = (f'<svg style="position:absolute;top:0;left:0;width:100%;height:100%;overflow:visible;" '
           f'viewBox="0 0 {viewBox_w} {viewBox_h}" preserveAspectRatio="none">'
           f'<path d="{path_data}" fill="{fill_attr}" fill-rule="evenodd" stroke="{stroke_attr}" stroke-width="{stroke_w}" stroke-linejoin="round" />'
           f'</svg>')
    return svg


def parse_preset_geom(prst, ext_cx, ext_cy, fill_color, line_color, line_w_pt, no_fill, no_line):
    w = int(ext_cx) if ext_cx else 1
    h = int(ext_cy) if ext_cy else 1

    fill_attr = 'none' if no_fill else (fill_color or 'none')
    stroke_attr = 'none' if no_line else (line_color or 'none')
    stroke_w = f'{line_w_pt}pt' if line_w_pt and not no_line else '0'

    d = ''
    use_ellipse = False

    if prst == 'rect':
        d = f'M 0 0 L {w} 0 L {w} {h} L 0 {h} Z'
    elif prst == 'line':
        # Word 的"直接连接符"在 PDF 导出时被渲染为"填充细矩形"（fill=线颜色），
        # 而不是描边路径。若按 SVG <path stroke> 渲染，在 viewBox+preserveAspectRatio="none"
        # 下 0.75pt 笔画会被纵向压扁到几乎不可见（这就是"横线不显示"的原因）。
        # 改为输出 <rect fill=线色>，与 PDF 矢量事实完全一致；细长矩形天然就是可见的实心条。
        rect_fill = line_color or fill_attr
        svg = (f'<svg style="position:absolute;top:0;left:0;width:100%;height:100%;overflow:visible;" '
               f'viewBox="0 0 {w} {h}" preserveAspectRatio="none">'
               f'<rect x="0" y="0" width="{w}" height="{h}" fill="{rect_fill}" />'
               f'</svg>')
        return svg
    elif prst == 'rtTriangle':
        # OOXML 标准 preset（ECMA-376）的 rtTriangle：直角在左下角 (l,b)。
        #   顶点顺序：moveTo(l,b) → lnTo(l,t) → lnTo(r,b)，即 (0,h) 左下、(0,0) 左上、(w,h) 右下。
        # 旋转(rot)/翻转(flipH/flipV) 由下方通用变换统一施加，这里只写标准几何，
        # 从而保证对任意文档、任意旋转角都能按 OOXML 规范正确还原。
        # 本仓库实测：docx 中 6 个 rtTriangle 均带 rot=180°(无 flip)；
        # 用此“直角在左下”的标准几何 + 180° 旋转，直角恰好落在右上，与 PDF 原版矢量顶点完全一致。
        # （早期写成“直角在左上”的版本，转 180° 后直角落在右下，与 PDF 相比整体上下翻转，方向错误。）
        d = f'M 0 0 L 0 {h} L {w} {h} Z'
    elif prst == 'triangle':
        d = f'M {w//2} 0 L {w} {h} L 0 {h} Z'
    elif prst == 'ellipse':
        use_ellipse = True
    elif prst == 'roundRect':
        r = min(w, h) // 5
        d = (f'M {r} 0 L {w-r} 0 Q {w} 0 {w} {r} L {w} {h-r} Q {w} {h} {w-r} {h} '
             f'L {r} {h} Q 0 {h} 0 {h-r} L 0 {r} Q 0 0 {r} 0 Z')
    elif prst == 'diamond':
        d = f'M {w//2} 0 L {w} {h//2} L {w//2} {h} L 0 {h//2} Z'
    elif prst == 'pentagon':
        d = f'M {w//2} 0 L {w} {h//2} L {int(w*0.8)} {h} L {int(w*0.2)} {h} L 0 {h//2} Z'
    elif prst == 'hexagon':
        d = (f'M {int(w*0.25)} 0 L {int(w*0.75)} 0 L {w} {h//2} '
             f'L {int(w*0.75)} {h} L {int(w*0.25)} {h} L 0 {h//2} Z')
    elif prst == 'star5':
        cx, cy = w//2, h//2
        outer_r = min(w, h) // 2
        inner_r = int(outer_r * 0.4)
        pts = []
        for i in range(10):
            angle = math.radians(-90 + i * 36)
            r = outer_r if i % 2 == 0 else inner_r
            x = int(cx + r * math.cos(angle))
            y = int(cy + r * math.sin(angle))
            pts.append(f'{x} {y}')
        d = 'M ' + ' L '.join(pts) + ' Z'
    elif prst == 'arrow':
        aw = int(w * 0.3)
        ah = int(h * 0.3)
        d = (f'M 0 {h//2 - ah//2} L {w-aw} {h//2 - ah//2} L {w-aw} 0 '
             f'L {w} {h//2} L {w-aw} {h} L {w-aw} {h//2 + ah//2} L 0 {h//2 + ah//2} Z')
    else:
        # 默认按矩形处理
        d = f'M 0 0 L {w} 0 L {w} {h} L 0 {h} Z'

    if use_ellipse:
        rx = w / 2
        ry = h / 2
        cx = w / 2
        cy = h / 2
        svg = (f'<svg style="position:absolute;top:0;left:0;width:100%;height:100%;overflow:visible;" '
               f'viewBox="0 0 {w} {h}" preserveAspectRatio="none">'
               f'<ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" fill="{fill_attr}" stroke="{stroke_attr}" stroke-width="{stroke_w}" />'
               f'</svg>')
    else:
        svg = (f'<svg style="position:absolute;top:0;left:0;width:100%;height:100%;overflow:visible;" '
               f'viewBox="0 0 {w} {h}" preserveAspectRatio="none">'
               f'<path d="{d}" fill="{fill_attr}" stroke="{stroke_attr}" stroke-width="{stroke_w}" stroke-linejoin="round" />'
               f'</svg>')
    return svg


# ===================================================================
#  组合(嵌套组)坐标变换
# ===================================================================
def parse_group_transform(node):
    """读取一个组合节点(wgp/grpSp)的 grpSpPr，返回其本地坐标变换。

    返回 (off_x, off_y, chOff_x, chOff_y, chExt_x, chExt_y, sx, sy)。
    子坐标 c 经该变换映射到父坐标：d = off + (c - chOff) * (sx, sy)，
    其中真实的组缩放为 sx = ext.x / chExt.x, sy = ext.y / chExt.y。

    注意：缩放必须用 chExt（子坐标空间尺寸）作分母，绝不能用 chOff
    （chOff 只是子坐标原点的偏移量，通常极小，误用它会得到荒谬的放大倍数）。"""
    if node is None:
        return 0, 0, 0, 0, 0, 0, 1.0, 1.0
    grpSpPr = None
    for c in node:
        if isinstance(c.tag, str) and c.tag.split('}')[-1] == 'grpSpPr':
            grpSpPr = c
            break
    if grpSpPr is None:
        return 0, 0, 0, 0, 0, 0, 1.0, 1.0
    xf = grpSpPr.find(f'{{{A}}}xfrm')
    if xf is None:
        return 0, 0, 0, 0, 0, 0, 1.0, 1.0
    off = xf.find(f'{{{A}}}off'); ext = xf.find(f'{{{A}}}ext')
    chOff = xf.find(f'{{{A}}}chOff'); chExt = xf.find(f'{{{A}}}chExt')
    gox = int(off.get('x') or 0) if off is not None else 0
    goy = int(off.get('y') or 0) if off is not None else 0
    gex = int(ext.get('cx') or 0) if ext is not None else 0
    gey = int(ext.get('cy') or 0) if ext is not None else 0
    gchox = int(chOff.get('x') or 0) if chOff is not None else 0
    gchoy = int(chOff.get('y') or 0) if chOff is not None else 0
    gchex = int(chExt.get('cx') or 0) if chExt is not None else 0
    gchey = int(chExt.get('cy') or 0) if chExt is not None else 0
    gsx = (gex / gchex) if gchex else 1.0
    gsy = (gey / gchey) if gchey else 1.0
    return gox, goy, gchox, gchoy, gchex, gchey, gsx, gsy


def render_group(node, tx, ty, sx, sy, fallback_cx, fallback_cy,
                 all_elements, rel_height, behind_doc):
    """递归渲染组合节点，正确叠加各级嵌套组的 chOff/chExt 坐标缩放。
    tx/ty/sx/sy 为该节点"之前"已累积的变换，使得：
        absolute = (tx, ty) + child_coord * (sx, sy)。"""
    gox, goy, gchox, gchoy, gchex, gchey, gsx, gsy = parse_group_transform(node)
    nsx = sx * gsx
    nsy = sy * gsy
    ntx = tx + (gox - gchox * gsx) * sx
    nty = ty + (goy - gchoy * gsy) * sy
    for child in list(node):
        lname = child.tag.split('}')[-1] if isinstance(child.tag, str) else ''
        if lname == 'wsp':
            html = parse_shape(child, ntx, nty, nsx, nsy,
                               fallback_cx, fallback_cy, is_direct_shape=False)
            if html:
                all_elements.append((rel_height, behind_doc == '1', html))
        elif lname in ('wgp', 'grpSp'):
            render_group(child, ntx, nty, nsx, nsy,
                         fallback_cx, fallback_cy, all_elements, rel_height, behind_doc)


# ===================================================================
#  形状解析
# ===================================================================
def parse_shape(wsp, tx, ty, sx, sy, fallback_cx, fallback_cy,
                is_direct_shape=False):
    if wsp is None:
        return ''

    cnvpr = gchild(wsp, 'cNvPr', 'wps')
    shape_name = gattr(cnvpr, 'name', 'wps') if cnvpr is not None else ''
    shape_id = gattr(cnvpr, 'id', 'wps') if cnvpr is not None else '0'

    sppr = gchild(wsp, 'spPr', 'wps')
    if sppr is None:
        sppr = gchild(wsp, 'spPr', 'a')

    xfrm = gchild(sppr, 'xfrm', 'a') if sppr is not None else None

    child_off_x = 0
    child_off_y = 0
    child_ext_cx = 0
    child_ext_cy = 0
    rot = 0
    flipH = False
    flipV = False

    if xfrm is not None:
        off = gchild(xfrm, 'off', 'a')
        ext = gchild(xfrm, 'ext', 'a')
        if off is not None:
            raw_off_x = int(gattr(off, 'x', 'a') or '0')
            raw_off_y = int(gattr(off, 'y', 'a') or '0')

            if is_direct_shape and ext is not None:
                ext_cx = int(gattr(ext, 'cx', 'a') or '0')
                ext_cy = int(gattr(ext, 'cy', 'a') or '0')
                if ext_cx == fallback_cx and ext_cy == fallback_cy:
                    child_off_x = 0
                    child_off_y = 0
                else:
                    child_off_x = raw_off_x
                    child_off_y = raw_off_y
            else:
                child_off_x = raw_off_x
                child_off_y = raw_off_y
        if ext is not None:
            child_ext_cx = int(gattr(ext, 'cx', 'a') or '0')
            child_ext_cy = int(gattr(ext, 'cy', 'a') or '0')
        rot_val = gattr(xfrm, 'rot', 'a') or '0'
        rot = int(rot_val) / 60000 if rot_val else 0
        flipH = gattr(xfrm, 'flipH', 'a') == '1'
        flipV = gattr(xfrm, 'flipV', 'a') == '1'

    prst_geom = gchild(sppr, 'prstGeom', 'a') if sppr is not None else None
    cust_geom = gchild(sppr, 'custGeom', 'a') if sppr is not None else None

    # 填充
    fill_color = None
    no_fill = False
    if sppr is not None:
        if gchild(sppr, 'noFill', 'a') is not None:
            no_fill = True
        else:
            solid_fill = gchild(sppr, 'solidFill', 'a')
            if solid_fill is not None:
                srgb = gchild(solid_fill, 'srgbClr', 'a')
                scheme = gchild(solid_fill, 'schemeClr', 'a')
                if srgb is not None:
                    fill_color = '#' + (gattr(srgb, 'val', 'a') or '000000').upper()
                    alpha_elem = gchild(srgb, 'alpha', 'a')
                    if alpha_elem is not None:
                        alpha_val = int(gattr(alpha_elem, 'val', 'a') or '100000') / 100000
                        fill_color = hex_to_rgba(fill_color, alpha_val)
                elif scheme is not None:
                    fill_color = parse_scheme_clr(scheme)

    # 渐变填充（简化处理：取第一个色标）
    if fill_color is None and sppr is not None:
        grad_fill = gchild(sppr, 'gradFill', 'a')
        if grad_fill is not None:
            gs_lst = gchild(grad_fill, 'gsLst', 'a')
            if gs_lst is not None:
                gs_elems = gchildren(gs_lst, 'gs', 'a')
                if gs_elems:
                    first_gs = gs_elems[0]
                    srgb = gchild(first_gs, 'srgbClr', 'a')
                    if srgb is not None:
                        fill_color = '#' + (gattr(srgb, 'val', 'a') or '000000').upper()

    # 线条
    line_color = None
    no_line = False
    line_w_pt = 0
    if sppr is not None:
        ln = gchild(sppr, 'ln', 'a')
        if ln is not None:
            w_val = gattr(ln, 'w', 'a')
            if w_val:
                line_w_pt = int(w_val) / EMU_PER_PT
            fill = gchild(ln, 'solidFill', 'a')
            if fill is not None:
                srgb = gchild(fill, 'srgbClr', 'a')
                scheme = gchild(fill, 'schemeClr', 'a')
                if srgb is not None:
                    line_color = '#' + (gattr(srgb, 'val', 'a') or '000000').upper()
                elif scheme is not None:
                    line_color = parse_scheme_clr(scheme)
            else:
                if gchild(ln, 'noFill', 'a') is not None:
                    no_line = True
                else:
                    # 有轮廓但未指定颜色 → OOXML 默认黑色描边
                    line_color = '#000000'
            # OOXML 默认线宽 0.75pt（<a:ln> 未写 w 时）
            if line_w_pt == 0 and not no_line:
                line_w_pt = 0.75
        else:
            no_line = True

    # 文本框内容
    txbx = gchild(wsp, 'txbx', 'wps')
    txbx_content_html = ''
    if txbx is not None:
        txbx_content = gchild(txbx, 'txbxContent', 'w')
        txbx_content_html = parse_txbx_content(txbx_content)

    abs_x_emu = tx + child_off_x * sx
    abs_y_emu = ty + child_off_y * sy
    is_line = prst_geom is not None and gattr(prst_geom, 'prst', 'a') == 'line'
    if is_line:
        # 线形：直接使用自身 ext（某一维为 0 表示一维线），不要回退到组尺寸
        w_emu = child_ext_cx * sx
        h_emu = child_ext_cy * sy
    else:
        w_emu = child_ext_cx * sx if child_ext_cx > 0 else fallback_cx
        h_emu = child_ext_cy * sy if child_ext_cy > 0 else fallback_cy

    x_mm = emu_to_mm(abs_x_emu)
    y_mm = emu_to_mm(abs_y_emu)
    w_mm = emu_to_mm(w_emu)
    h_mm = emu_to_mm(h_emu)

    # 一维线（水平/垂直）：细边为 0 时至少保留描边宽度，否则线不可见且会变成斜线
    sw_emu = int(line_w_pt * EMU_PER_PT) if (line_w_pt and not no_line) else 0
    if is_line and sw_emu > 0:
        if w_emu == 0 and h_emu == 0:
            w_emu = h_emu = sw_emu
        elif w_emu == 0:
            w_emu = sw_emu
            x_mm -= emu_to_mm(sw_emu) / 2.0
        elif h_emu == 0:
            h_emu = sw_emu
            y_mm -= emu_to_mm(sw_emu) / 2.0
        w_mm = emu_to_mm(w_emu)
        h_mm = emu_to_mm(h_emu)

    transform_parts = []
    if rot:
        transform_parts.append(f'rotate({rot}deg)')
    if flipH or flipV:
        scale_x = -1 if flipH else 1
        scale_y = -1 if flipV else 1
        transform_parts.append(f'scale({scale_x}, {scale_y})')

    transform_style = f' transform: {" ".join(transform_parts)};' if transform_parts else ''
    transform_origin = ' transform-origin: center;' if transform_parts else ''

    div_style = f'position: absolute; left: {x_mm}mm; top: {y_mm}mm; width: {w_mm}mm; height: {h_mm}mm;{transform_style}{transform_origin}'

    inner_html = ''

    if cust_geom is not None:
        svg = parse_cust_geom_to_svg(cust_geom, fill_color, line_color, line_w_pt, no_fill, no_line)
        if svg:
            inner_html += svg
    elif prst_geom is not None:
        prst = gattr(prst_geom, 'prst', 'a')
        svg = parse_preset_geom(prst, w_emu, h_emu, fill_color, line_color, line_w_pt, no_fill, no_line)
        if svg:
            inner_html += svg

    if txbx_content_html:
        body_pr = gchild(wsp, 'bodyPr', 'wps')
        anchor_val = 't'
        if body_pr is not None:
            anchor_attr = gattr(body_pr, 'anchor', 'wps')
            if anchor_attr:
                anchor_val = anchor_attr

        text_align_style = ''
        if anchor_val == 'ctr':
            text_align_style = ' display: flex; flex-direction: column; justify-content: center;'
        elif anchor_val == 'b':
            text_align_style = ' display: flex; flex-direction: column; justify-content: flex-end;'

        l_ins = r_ins = t_ins = b_ins = 91440
        if body_pr is not None:
            lIns = gattr(body_pr, 'lIns', 'wps')
            rIns = gattr(body_pr, 'rIns', 'wps')
            tIns = gattr(body_pr, 'tIns', 'wps')
            bIns = gattr(body_pr, 'bIns', 'wps')
            if lIns: l_ins = int(lIns)
            if rIns: r_ins = int(rIns)
            if tIns: t_ins = int(tIns)
            if bIns: b_ins = int(bIns)

        pad_l = emu_to_mm(l_ins)
        pad_r = emu_to_mm(r_ins)
        pad_t = emu_to_mm(t_ins)
        pad_b = emu_to_mm(b_ins)

        inner_html += (f'<div style="position: relative; z-index: 2; width: 100%; height: 100%; '
                       f'overflow: hidden; padding: {pad_t}mm {pad_r}mm {pad_b}mm {pad_l}mm;{text_align_style}">'
                       f'{txbx_content_html}</div>')

    if txbx_content_html and fill_color and not no_fill:
        elem_class = "docx-element label"
    elif txbx_content_html:
        elem_class = "docx-element text-box"
    elif shape_name and ('矩形' in shape_name or '多边形' in shape_name or '三角形' in shape_name):
        elem_class = "docx-element shape"
    elif shape_name and '连接符' in shape_name:
        elem_class = "docx-element connector"
    else:
        elem_class = "docx-element shape"

    safe_name = html_lib.escape(shape_name or '')

    return f'<div class="{elem_class}" style="{div_style}" data-name="{safe_name}" data-id="{shape_id}">{inner_html}</div>'


# ===================================================================
#  段落高度估算
# ===================================================================
def estimate_paragraph_height_emu(p_elem):
    """估算段落在页面中占据的高度 (EMU)"""
    ppr = gchild(p_elem, 'pPr', 'w')

    para_mark_sz_half_pt = 20
    if ppr is not None:
        rPr = gchild(ppr, 'rPr', 'w')
        if rPr is not None:
            sz = gchild(rPr, 'sz', 'w')
            if sz is not None:
                val = gattr(sz, 'val', 'w')
                if val:
                    para_mark_sz_half_pt = int(val)

    font_pt = para_mark_sz_half_pt / 2.0

    line_rule = 'auto'
    line_val = ''
    after_twips = 0
    before_twips = 0

    if ppr is not None:
        spacing = gchild(ppr, 'spacing', 'w')
        if spacing is not None:
            before_twips = int(gattr(spacing, 'before', 'w') or '0')
            after_twips = int(gattr(spacing, 'after', 'w') or '0')
            line_val = gattr(spacing, 'line', 'w') or ''
            line_rule = gattr(spacing, 'lineRule', 'w') or 'auto'

    if line_val and line_rule == 'exact':
        line_height_pt = int(line_val) / 20.0
    elif line_val and line_rule == 'atLeast':
        line_height_pt = max(int(line_val) / 20.0, font_pt)
    elif line_val and line_rule == 'auto':
        line_multiple = int(line_val) / 240.0
        line_height_pt = font_pt * line_multiple
    else:
        line_height_pt = font_pt

    before_pt = before_twips / 20.0
    after_pt = after_twips / 20.0

    total_pt = line_height_pt + before_pt + after_pt
    total_emu = int(total_pt * EMU_PER_PT)

    return total_emu


# ===================================================================
#  表格解析
# ===================================================================
def parse_table(tbl_elem, current_y_emu, mar_emu, page_w_emu, mar_top_emu, rels, image_data):
    """解析 w:tbl 表格，返回 HTML + 估算高度"""
    if tbl_elem is None:
        return '', 0

    tbl_pr = gchild(tbl_elem, 'tblPr', 'w')
    tbl_w = 0
    if tbl_pr is not None:
        tbl_w_elem = gchild(tbl_pr, 'tblW', 'w')
        if tbl_w_elem is not None:
            w_val = gattr(tbl_w_elem, 'w', 'w')
            w_type = gattr(tbl_w_elem, 'type', 'w')
            if w_val:
                if w_type == 'pct':
                    tbl_w = int(int(w_val) / 5000 * page_w_emu)
                elif w_type == 'dxa':
                    tbl_w = int(w_val) * 635
                else:
                    tbl_w = int(w_val)

    if tbl_w == 0:
        tbl_w = page_w_emu - 2 * mar_emu

    # 单元格默认边距（tblCellMar），用于校正单元格内绝对定位浮动元素的横向参考点。
    # Word 的 anchor offset 是相对「单元格内容区」（已扣除 left/right margin）的，
    # 而 HTML position:absolute 是相对 td 的 padding 盒（此处 td 无 padding），
    # 故需把单元格 left/top margin 补偿进浮动元素的偏移，照片才不会整体偏左/偏上。
    # 若表格未显式定义 tblCellMar，回退 Word 默认值 108 twips (≈1.9mm)。
    cell_margin_l_emu = 108 * 635
    cell_margin_t_emu = 0
    if tbl_pr is not None:
        tcm = gchild(tbl_pr, 'tblCellMar', 'w')
        if tcm is not None:
            for side in ('left', 'top'):
                m = gchild(tcm, side, 'w')
                if m is not None:
                    val = gattr(m, 'w', 'w')
                    if val:
                        emu = int(val) * 635
                        if side == 'left':
                            cell_margin_l_emu = emu
                        else:
                            cell_margin_t_emu = emu

    # 收集行
    rows = gchildren(tbl_elem, 'tr', 'w')
    if not rows:
        return '', 0

    # OOXML 表格默认 tblLayout="fixed"，对应 HTML 必须显式 table-layout: fixed，
    # 否则浏览器按 table-layout: auto 把列宽按内容自动分配，我们设的 width: 13.5% 等
    # 百分比会被忽略，导致列被压扁、内容竖排（"出生年月"被拆成 出/生/年/月）。
    #
    # 进一步：固定布局下，列结构由第一行的 td 决定，但表头行（如"个人简历"）常常
    # 只有 1 个 colspan 单元格（width:100%），若按它定列则整表退化成 1 列，其他行的
    # 多列内容会被压缩。正确做法是从 docx 的 <w:tblGrid> 读出每一列的真实宽度
    # （gridCol w="dxa"），转换为百分比后输出为 <colgroup><col>——这样无论哪一行
    # 先出现，列结构都是 5 列；再叠加 table-layout: fixed，单元格百分比才会被尊重。
    grid_elem = gchild(tbl_elem, 'tblGrid', 'w')
    grid_widths_pct = []
    if grid_elem is not None:
        cols = gchildren(grid_elem, 'gridCol', 'w')
        grid_dxa = [int(gattr(c, 'w', 'w') or 0) for c in cols]
        total_dxa = sum(grid_dxa) or tbl_w
        grid_widths_pct = [w / total_dxa * 100 for w in grid_dxa]

    colgroup_html = '<colgroup>'
    for w_pct in grid_widths_pct:
        colgroup_html += f'<col style="width: {w_pct:.2f}%;">'
    colgroup_html += '</colgroup>'

    table_html = (f'<table style="border-collapse: collapse; width: 100%; '
                  f'table-layout: fixed;">{colgroup_html}')
    total_height_emu = 0

    float_elements = []  # 保留返回兼容（单元格内浮动现已注入 <td>）

    for tr in rows:
        tr_height = 0
        tr_pr = gchild(tr, 'trPr', 'w')
        if tr_pr is not None:
            tr_h_elem = gchild(tr_pr, 'trHeight', 'w')
            if tr_h_elem is not None:
                h_val = gattr(tr_h_elem, 'val', 'w')
                if h_val:
                    tr_height = int(h_val) * 635

        # 不为 trHeight=0 的行设默认值：保持 0，让浏览器按内容自动算行高，
        # 否则每行都被强制一个固定高度，整体版式就会错乱。
        tr_style = f' style="height: {emu_to_mm(tr_height):.2f}mm;"' if tr_height > 0 else ''
        table_html += f'<tr{tr_style}>'
        cells = gchildren(tr, 'tc', 'w')
        for tc in cells:
            tc_pr = gchild(tc, 'tcPr', 'w')
            tc_w = 0
            tc_grid_span = 1
            if tc_pr is not None:
                tc_w_elem = gchild(tc_pr, 'tcW', 'w')
                if tc_w_elem is not None:
                    w_val = gattr(tc_w_elem, 'w', 'w')
                    w_type = gattr(tc_w_elem, 'type', 'w')
                    if w_val:
                        if w_type == 'pct':
                            tc_w = int(int(w_val) / 5000 * tbl_w)
                        else:
                            tc_w = int(w_val) * 635
                gs_elem = gchild(tc_pr, 'gridSpan', 'w')
                if gs_elem is not None:
                    gs_val = gattr(gs_elem, 'val', 'w')
                    if gs_val:
                        tc_grid_span = int(gs_val)

            # 单元格背景
            bg_color = ''
            if tc_pr is not None:
                shd = gchild(tc_pr, 'shd', 'w')
                if shd is not None:
                    fill_val = gattr(shd, 'fill', 'w')
                    if fill_val and fill_val != 'auto':
                        bg_color = f' background-color: #{fill_val.upper()};'

            # 边框
            borders_style = ''
            if tc_pr is not None:
                tc_borders = gchild(tc_pr, 'tcBorders', 'w')
                if tc_borders is not None:
                    for side in ['top', 'left', 'bottom', 'right']:
                        border = gchild(tc_borders, side, 'w')
                        if border is not None:
                            b_val = gattr(border, 'val', 'w')
                            b_sz = gattr(border, 'sz', 'w')
                            b_color = gattr(border, 'color', 'w')
                            if b_val and b_val != 'nil':
                                b_sz_pt = int(b_sz or '4') / 8.0
                                b_color_val = f'#{b_color.upper()}' if b_color and b_color != 'auto' else '#000000'
                                borders_style += f' border-{side}: {b_sz_pt}pt solid {b_color_val};'

            # 解析单元格内容（文字）
            cell_html = ''
            for p in gchildren(tc, 'p', 'w'):
                p_html = parse_paragraph(p)
                if p_html:
                    cell_html += p_html

            # 单元格内浮动元素：锚定在单元格里的 wp:anchor（图片/形状）。
            # 例：简历模板右上角人像照片就是「固定在表格单元格内的浮动图片」，
            # 之前 parse_table 只渲文字、漏掉 anchor，导致照片整张丢失。
            #
            # 定位策略：不按「页面绝对坐标」估算（表格行高由 CSS 决定，EMU 估算的
            # 行顶会和真实渲染错位十几毫米）。改为把元素注入所在 <td> 内部、用
            # position:absolute 相对该单元格定位，由浏览器按真实单元格位置渲染，
            # 纵向天然对齐；横向偏移用 OOXML 的 column 偏移量（相对列左缘=单元格左缘）。
            cell_anchors_html = ''
            for anchor in tc.findall(f'.//{{{WP}}}anchor'):
                cell_anchors_html += render_cell_anchor(
                    anchor, rels, image_data, cell_margin_l_emu, cell_margin_t_emu)

            w_pct = (tc_w / tbl_w * 100) if tbl_w > 0 else 100
            colspan_attr = f' colspan="{tc_grid_span}"' if tc_grid_span > 1 else ''
            td_style = (f'position: relative; width: {w_pct:.1f}%;'
                        f'{bg_color}{borders_style} vertical-align: top;')
            table_html += f'<td{colspan_attr} style="{td_style}">{cell_html}{cell_anchors_html}</td>'

        table_html += '</tr>'

    table_html += '</table>'

    y_mm = emu_to_mm(current_y_emu)
    w_mm = emu_to_mm(tbl_w)
    h_mm = emu_to_mm(total_height_emu)

    div_html = (f'<div class="docx-element table" '
                f'style="position: absolute; left: {emu_to_mm(mar_emu)}mm; top: {y_mm}mm; width: {w_mm}mm;">'
                f'{table_html}</div>')

    return div_html, total_height_emu, float_elements


def get_default_font_size_pt(docx_dir):
    """读取文档默认字号（半磅→pt）。

    OOXML 里正文大多 run 不写 w:sz，而是继承 Normal 样式的字号。
    之前转换器不给这些 run 设 font-size，浏览器就用默认 16px，导致整页偏大偏黑。
    这里从 styles.xml 读出默认字号，交给 body 作为基准，无显式字号的 run 自动继承。

    优先级：默认段落样式(w:default=1)的 w:sz > docDefaults/rPrDefault 的 w:sz > 10.5pt。
    """
    styles_path = os.path.join(docx_dir, 'word', 'styles.xml')
    if not os.path.exists(styles_path):
        return 10.5
    try:
        st = ET.parse(styles_path).getroot()
        # 1) 默认段落样式的 sz
        for st_node in st.iter(f'{{{W}}}style'):
            if gattr(st_node, 'default', 'w') == '1' and gattr(st_node, 'type', 'w') == 'paragraph':
                sz = st_node.find(f'.//{{{W}}}sz')
                if sz is not None:
                    return int(gattr(sz, 'val', 'w') or 21) / 2.0
        # 2) docDefaults / rPrDefault 的 sz
        dd = st.find(f'.//{{{W}}}rPrDefault/{{{W}}}rPr')
        if dd is not None:
            sz = dd.find(f'{{{W}}}sz')
            if sz is not None:
                return int(gattr(sz, 'val', 'w') or 21) / 2.0
    except Exception:
        pass
    return 10.5


def render_anchor(anchor, current_para_y, mar_left_emu, mar_top_emu, rels, image_data):
    """渲染一个 wp:anchor 浮动元素（图片 / 形状 / 组合）。

    返回 [(rel_height, behind_doc, html_str)]，坐标以「页面绝对 EMU」给出，
    与顶层 body 浮动元素完全一致，可直接 append 到 all_elements / body_content_elements。

    mar_left_emu 在此语义为「锚定参考列的左缘页面坐标」：
      - 顶层段落里的 anchor：传 mar_left_emu（页面左边距），relativeFrom=column 即相对页面左缘；
      - 表格单元格里的 anchor：传「该单元格所在列的左缘页面坐标」，relativeFrom=column
        仍能正确换算到页面绝对坐标。
    """
    elements = []

    ph = gchild(anchor, 'positionH', 'wp')
    pv = gchild(anchor, 'positionV', 'wp')

    h_relative = gattr(ph, 'relativeFrom', 'wp') if ph is not None else 'column'
    v_relative = gattr(pv, 'relativeFrom', 'wp') if pv is not None else 'paragraph'

    h_offset = 0
    v_offset = 0
    if ph is not None:
        posoff = gchild(ph, 'posOffset', 'wp')
        if posoff is not None:
            h_offset = int(posoff.text)
    if pv is not None:
        posoff = gchild(pv, 'posOffset', 'wp')
        if posoff is not None:
            v_offset = int(posoff.text)

    extent = gchild(anchor, 'extent', 'wp')
    cx = int(gattr(extent, 'cx', 'wp') or '0') if extent is not None else 0
    cy = int(gattr(extent, 'cy', 'wp') or '0') if extent is not None else 0

    behind_doc = gattr(anchor, 'behindDoc', 'wp') or '0'
    rel_height = int(gattr(anchor, 'relativeHeight', 'wp') or '0')

    # 浮动参考点（页面绝对 EMU）
    if h_relative == 'page':
        real_h = h_offset
    elif h_relative == 'margin':
        real_h = mar_left_emu + h_offset
    else:  # column / character / paragraph：相对「列左缘」（此处 mar_left_emu 已是列左缘）
        real_h = mar_left_emu + h_offset

    if v_relative == 'page':
        real_v = v_offset
    elif v_relative == 'margin':
        real_v = mar_top_emu + v_offset
    else:  # paragraph：相对段落顶
        real_v = current_para_y + v_offset

    # 组合 / 形状 / 图片
    wpg = None
    for e in anchor.iter():
        if isinstance(e.tag, str) and e.tag.split('}')[-1] == 'wgp':
            wpg = e
            break

    if wpg is not None:
        render_group(wpg, real_h, real_v, 1.0, 1.0, cx, cy,
                     elements, rel_height, behind_doc)
    else:
        wsps = anchor.findall(f'.//{{{WPS}}}wsp')
        for wsp in wsps:
            html_str = parse_shape(wsp, real_h, real_v, 1.0, 1.0, cx, cy, is_direct_shape=True)
            if html_str:
                elements.append((rel_height, behind_doc == '1', html_str))

        pics = anchor.findall(f'.//{{{NS["pic"]}}}pic')
        for pic_elem in pics:
            blip = pic_elem.find(f'.//{{{A}}}blip')
            if blip is not None:
                embed = blip.get(f'{{{NS["r"]}}}embed', '')
                target = rels.get(embed, '')
                if target.startswith('media/'):
                    img_file = target.split('/')[-1]
                    if img_file in image_data:
                        img_src = image_data[img_file]

                        sp_pr = gchild(pic_elem, 'spPr', 'pic')
                        pic_xfrm = gchild(sp_pr, 'xfrm', 'a') if sp_pr is not None else None

                        if pic_xfrm is not None:
                            off = gchild(pic_xfrm, 'off', 'a')
                            ext = gchild(pic_xfrm, 'ext', 'a')
                            if off is not None and ext is not None:
                                pic_x = int(gattr(off, 'x', 'a') or '0')
                                pic_y = int(gattr(off, 'y', 'a') or '0')
                                pic_cx = int(gattr(ext, 'cx', 'a') or '0')
                                pic_cy = int(gattr(ext, 'cy', 'a') or '0')
                            else:
                                pic_x = pic_y = 0
                                pic_cx = cx
                                pic_cy = cy
                        else:
                            pic_x = pic_y = 0
                            pic_cx = cx
                            pic_cy = cy

                        abs_x = real_h + pic_x
                        abs_y = real_v + pic_y

                        x_mm = emu_to_mm(abs_x)
                        y_mm = emu_to_mm(abs_y)
                        w_mm = emu_to_mm(pic_cx)
                        h_mm = emu_to_mm(pic_cy)

                        z_idx = rel_height - 251640000
                        html_str = (f'<div class="docx-element image" '
                                    f'style="position: absolute; left: {x_mm}mm; top: {y_mm}mm; '
                                    f'width: {w_mm}mm; height: {h_mm}mm; z-index: {z_idx};">'
                                    f'<img src="{img_src}" style="width:100%;height:100%;object-fit:fill;" />'
                                    f'</div>')
                        elements.append((rel_height, behind_doc == '1', html_str))

    return elements


def render_cell_anchor(anchor, rels, image_data, cell_margin_l_emu=0, cell_margin_t_emu=0):
    """渲染「锚定在表格单元格内」的浮动元素，返回 HTML 片段（相对所在 <td> 定位）。

    与 render_anchor（页面绝对坐标）不同：单元格内浮动元素若用页面坐标估算纵向位置，
    会因表格行高由 CSS 实际渲染、与 EMU 估算不完全一致而错位十几毫米。
    因此这里把元素注入 <td>（该 td 已设 position:relative），用 OOXML 的
    positionH/positionV 偏移量作为相对该单元格左上角的 CSS 偏移，由浏览器按真实
    单元格位置渲染，纵向天然对齐。

    适用前提：单元格内 anchor 一般用 relativeFrom=column（横向，相对列左缘=单元格左缘）
    与 relativeFrom=paragraph（纵向，相对段落顶≈单元格内容顶）。

    cell_margin_l_emu / cell_margin_t_emu：单元格内容区 left/top 边距（EMU）。
    Word 的 anchor offset 是相对「内容区」（已扣除单元格 margin）给出的，而 HTML
    position:absolute 相对 td 的 padding 盒（本转换器未给 td 设 padding），因此要把
    单元格 left/top margin 补回偏移量，浮动照片才不会相对 PDF 整体偏左/偏上约 1.9mm。
    """
    out = []

    ph = gchild(anchor, 'positionH', 'wp')
    pv = gchild(anchor, 'positionV', 'wp')
    h_relative = gattr(ph, 'relativeFrom', 'wp') if ph is not None else 'column'
    v_relative = gattr(pv, 'relativeFrom', 'wp') if pv is not None else 'paragraph'

    h_offset = 0
    v_offset = 0
    if ph is not None:
        posoff = gchild(ph, 'posOffset', 'wp')
        if posoff is not None:
            h_offset = int(posoff.text)
    if pv is not None:
        posoff = gchild(pv, 'posOffset', 'wp')
        if posoff is not None:
            v_offset = int(posoff.text)

    # 补偿单元格内容区边距：Word anchor offset 相对内容区（已扣 margin），
    # HTML 绝对定位相对 td padding 盒（无 padding），故加回 left/top margin。
    h_offset += cell_margin_l_emu
    v_offset += cell_margin_t_emu

    extent = gchild(anchor, 'extent', 'wp')
    cx = int(gattr(extent, 'cx', 'wp') or '0') if extent is not None else 0
    cy = int(gattr(extent, 'cy', 'wp') or '0') if extent is not None else 0
    behind_doc = gattr(anchor, 'behindDoc', 'wp') or '0'

    # 形状
    for wsp in anchor.findall(f'.//{{{WPS}}}wsp'):
        html_str = parse_shape(wsp, 0, 0, 1.0, 1.0, cx, cy, is_direct_shape=True)
        if html_str:
            out.append(_wrap_cell_anchor(html_str, h_offset, v_offset, cx, cy, behind_doc))

    # 图片
    for pic_elem in anchor.findall(f'.//{{{NS["pic"]}}}pic'):
        blip = pic_elem.find(f'.//{{{A}}}blip')
        if blip is not None:
            embed = blip.get(f'{{{NS["r"]}}}embed', '')
            target = rels.get(embed, '')
            if target.startswith('media/'):
                img_file = target.split('/')[-1]
                if img_file in image_data:
                    img_src = image_data[img_file]
                    sp_pr = gchild(pic_elem, 'spPr', 'pic')
                    pic_xfrm = gchild(sp_pr, 'xfrm', 'a') if sp_pr is not None else None
                    if pic_xfrm is not None:
                        off = gchild(pic_xfrm, 'off', 'a')
                        ext = gchild(pic_xfrm, 'ext', 'a')
                        if off is not None and ext is not None:
                            pic_x = int(gattr(off, 'x', 'a') or '0')
                            pic_y = int(gattr(off, 'y', 'a') or '0')
                            pic_cx = int(gattr(ext, 'cx', 'a') or '0')
                            pic_cy = int(gattr(ext, 'cy', 'a') or '0')
                        else:
                            pic_x = pic_y = 0
                            pic_cx = cx
                            pic_cy = cy
                    else:
                        pic_x = pic_y = 0
                        pic_cx = cx
                        pic_cy = cy
                    left_emu = h_offset + pic_x
                    top_emu = v_offset + pic_y
                    w_mm = emu_to_mm(pic_cx)
                    h_mm = emu_to_mm(pic_cy)
                    l_mm = emu_to_mm(left_emu)
                    t_mm = emu_to_mm(top_emu)
                    z = 1000 if behind_doc != '1' else -1
                    img_html = (f'<div style="position: absolute; left: {l_mm}mm; top: {t_mm}mm; '
                                f'width: {w_mm}mm; height: {h_mm}mm; z-index: {z};">'
                                f'<img src="{img_src}" style="width:100%;height:100%;object-fit:fill;" /></div>')
                    out.append(img_html)
    return ''.join(out)


def _wrap_cell_anchor(inner_html, h_offset, v_offset, cx, cy, behind_doc):
    l_mm = emu_to_mm(h_offset)
    t_mm = emu_to_mm(v_offset)
    w_mm = emu_to_mm(cx)
    h_mm = emu_to_mm(cy)
    z = 1000 if behind_doc != '1' else -1
    return (f'<div style="position: absolute; left: {l_mm}mm; top: {t_mm}mm; '
            f'width: {w_mm}mm; height: {h_mm}mm; z-index: {z};">{inner_html}</div>')


# ===================================================================
#  主转换函数
# ===================================================================
def convert_docx_to_html(docx_path, output_path=None, debug=False):
    """
    将 .docx 文件转换为 1:1 还原的 HTML

    Args:
        docx_path: .docx 文件路径
        output_path: 输出 HTML 路径（默认同目录同名 .html）
        debug: 是否显示调试水印（默认关闭）
    """
    if not os.path.exists(docx_path):
        print(f'错误: 文件不存在: {docx_path}')
        return None

    if output_path is None:
        base = os.path.splitext(docx_path)[0]
        output_path = base + '.html'

    # 解压到临时目录
    temp_dir = tempfile.mkdtemp(prefix='docx2html_')
    try:
        print(f'正在解压: {docx_path}')
        with zipfile.ZipFile(docx_path, 'r') as z:
            z.extractall(temp_dir)

        return _do_convert(temp_dir, output_path, debug=debug)

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _do_convert(docx_dir, output_path, debug=False):
    """核心转换逻辑（接受已解压的目录）"""
    global THEME_COLORS

    # 加载主题颜色
    THEME_COLORS = load_theme_colors(docx_dir)
    print(f'主题颜色已加载: {len(THEME_COLORS)} 个')

    doc_xml_path = os.path.join(docx_dir, 'word', 'document.xml')
    if not os.path.exists(doc_xml_path):
        print(f'错误: 未找到 word/document.xml')
        return None

    tree = ET.parse(doc_xml_path)
    root = tree.getroot()
    body = root.find(f'{{{W}}}body')

    if body is None:
        print('错误: 未找到 body 元素')
        return None

    # 页面尺寸
    sectpr = body.find(f'{{{W}}}sectPr')
    page_w_twips = 11906
    page_h_twips = 16838
    mar_left_twips = 720
    mar_right_twips = 720
    mar_top_twips = 720
    mar_bottom_twips = 720

    if sectpr is not None:
        pgsz = sectpr.find(f'{{{W}}}pgSz')
        if pgsz is not None:
            page_w_twips = int(gattr(pgsz, 'w', 'w') or page_w_twips)
            page_h_twips = int(gattr(pgsz, 'h', 'w') or page_h_twips)
        pgmar = sectpr.find(f'{{{W}}}pgMar')
        if pgmar is not None:
            mar_left_twips = int(gattr(pgmar, 'left', 'w') or '720')
            mar_right_twips = int(gattr(pgmar, 'right', 'w') or '720')
            mar_top_twips = int(gattr(pgmar, 'top', 'w') or '720')
            mar_bottom_twips = int(gattr(pgmar, 'bottom', 'w') or '720')

    page_w_mm = page_w_twips / 1440 * 25.4
    page_h_mm = page_h_twips / 1440 * 25.4
    mar_left_mm = mar_left_twips / 1440 * 25.4
    mar_right_mm = mar_right_twips / 1440 * 25.4
    mar_top_mm = mar_top_twips / 1440 * 25.4
    mar_bottom_mm = mar_bottom_twips / 1440 * 25.4

    page_w_emu = int(page_w_mm * EMU_PER_MM)
    page_h_emu = int(page_h_mm * EMU_PER_MM)
    mar_left_emu = int(mar_left_mm * EMU_PER_MM)
    mar_right_emu = int(mar_right_mm * EMU_PER_MM)
    mar_top_emu = int(mar_top_mm * EMU_PER_MM)
    mar_bottom_emu = int(mar_bottom_mm * EMU_PER_MM)

    content_w_emu = page_w_emu - mar_left_emu - mar_right_emu
    content_h_emu = page_h_emu - mar_top_emu - mar_bottom_emu

    # 文档默认字号（半磅→pt）。无显式 w:sz 的正文 run 用它作基准，避免浏览器默认 16px 放大整页。
    default_font_pt = get_default_font_size_pt(docx_dir)

    # Word 正文段落（本模板每个段都显式 snapToGrid=0）不吸附基线网格，而是用 Word 的
    # “单倍行距”——约 1.36× 字号（微软雅黑 10.5pt 正文下，比浏览器 normal(~1.15×) 略松）。
    # 浏览器默认 normal 会让正文行偏挤、整页被压短，章节条/照片随之向上漂移。
    # 用无量纲系数（相对字号）设置 body 行高，可随各 run 字号安全缩放，不裁切大字号（如 30pt 横幅）。
    # 该校准值经与 PDF 真值逐带比对得到；若默认字体/模板差异较大可微调。
    body_line_factor = 1.37

    # 读取图片
    image_data = {}
    media_dir = os.path.join(docx_dir, 'word', 'media')
    if os.path.exists(media_dir):
        for f in os.listdir(media_dir):
            fpath = os.path.join(media_dir, f)
            if os.path.isfile(fpath):
                with open(fpath, 'rb') as img_f:
                    img_bytes = img_f.read()
                ext = os.path.splitext(f)[1].lower()
                mime_map = {
                    '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
                    '.png': 'image/png', '.gif': 'image/gif',
                    '.bmp': 'image/bmp', '.tif': 'image/tiff',
                    '.tiff': 'image/tiff', '.wmf': 'image/x-wmf',
                    '.emf': 'image/x-emf', '.svg': 'image/svg+xml',
                    '.webp': 'image/webp',
                }
                mime = mime_map.get(ext, 'image/png')
                image_data[f] = f'data:{mime};base64,{base64.b64encode(img_bytes).decode()}'
    print(f'图片已加载: {len(image_data)} 个')

    # 读取 relationships
    rels = {}
    rels_path = os.path.join(docx_dir, 'word', '_rels', 'document.xml.rels')
    if os.path.exists(rels_path):
        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        for rel in rels_root:
            rid = rel.get('Id', '')
            target = rel.get('Target', '')
            rels[rid] = target
    print(f'关系已加载: {len(rels)} 个')

    # 遍历 body 子元素
    body_children = list(body)

    all_elements = []  # (rel_height, behind_doc, html_str)
    body_content_elements = []  # 正文流式内容，按 Y 位置排列

    paragraph_y_emu = mar_top_emu

    for child in body_children:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            current_para_y = paragraph_y_emu
            p_height = estimate_paragraph_height_emu(child)

            # 检查段落是否有正文文字（非 anchor）
            has_body_text = False
            body_text_html = ''

            # 先处理 anchor（浮动元素）
            for anchor in child.findall(f'.//{{{WP}}}anchor'):
                ph = gchild(anchor, 'positionH', 'wp')
                pv = gchild(anchor, 'positionV', 'wp')

                h_relative = gattr(ph, 'relativeFrom', 'wp') if ph is not None else 'column'
                v_relative = gattr(pv, 'relativeFrom', 'wp') if pv is not None else 'paragraph'

                h_offset = 0
                v_offset = 0

                if ph is not None:
                    posoff = gchild(ph, 'posOffset', 'wp')
                    if posoff is not None:
                        h_offset = int(posoff.text)
                if pv is not None:
                    posoff = gchild(pv, 'posOffset', 'wp')
                    if posoff is not None:
                        v_offset = int(posoff.text)

                extent = gchild(anchor, 'extent', 'wp')
                cx = int(gattr(extent, 'cx', 'wp') or '0') if extent is not None else 0
                cy = int(gattr(extent, 'cy', 'wp') or '0') if extent is not None else 0

                behind_doc = gattr(anchor, 'behindDoc', 'wp') or '0'
                rel_height = int(gattr(anchor, 'relativeHeight', 'wp') or '0')

                # 计算实际位置
                if h_relative == 'page':
                    real_h = h_offset
                elif h_relative == 'margin':
                    real_h = mar_left_emu + h_offset
                else:
                    real_h = mar_left_emu + h_offset

                if v_relative == 'page':
                    real_v = v_offset
                elif v_relative == 'margin':
                    real_v = mar_top_emu + v_offset
                else:
                    real_v = current_para_y + v_offset

                # 区分 group 和 direct shape
                graphic = anchor.find(f'.//{{{A}}}graphic')
                uri = ''
                if graphic is not None:
                    graphic_data = graphic.find(f'{{{A}}}graphicData')
                    if graphic_data is not None:
                        uri = graphic_data.get('uri', '')

                # 组合：用 local-name 查找 wgp（避免命名空间前缀在 ElementTree.find 下的匹配问题）
                wpg = None
                for e in anchor.iter():
                    if isinstance(e.tag, str) and e.tag.split('}')[-1] == 'wgp':
                        wpg = e
                        break

                if wpg is not None:
                    # 组合（可能含多级嵌套）：递归叠加各组 chOff/chExt 坐标变换
                    render_group(wpg, real_h, real_v, 1.0, 1.0, cx, cy,
                                 all_elements, rel_height, behind_doc)
                else:
                    wsps = anchor.findall(f'.//{{{WPS}}}wsp')
                    for wsp in wsps:
                        html_str = parse_shape(wsp, real_h, real_v, 1.0, 1.0, cx, cy, is_direct_shape=True)
                        if html_str:
                            all_elements.append((rel_height, behind_doc == '1', html_str))

                    # 图片
                    pics = anchor.findall(f'.//{{{NS["pic"]}}}pic')
                    for pic_elem in pics:
                        blip = pic_elem.find(f'.//{{{A}}}blip')
                        if blip is not None:
                            embed = blip.get(f'{{{NS["r"]}}}embed', '')
                            target = rels.get(embed, '')
                            if target.startswith('media/'):
                                img_file = target.split('/')[-1]
                                if img_file in image_data:
                                    img_src = image_data[img_file]

                                    sp_pr = gchild(pic_elem, 'spPr', 'pic')
                                    pic_xfrm = gchild(sp_pr, 'xfrm', 'a') if sp_pr is not None else None

                                    if pic_xfrm is not None:
                                        off = gchild(pic_xfrm, 'off', 'a')
                                        ext = gchild(pic_xfrm, 'ext', 'a')
                                        if off is not None and ext is not None:
                                            pic_x = int(gattr(off, 'x', 'a') or '0')
                                            pic_y = int(gattr(off, 'y', 'a') or '0')
                                            pic_cx = int(gattr(ext, 'cx', 'a') or '0')
                                            pic_cy = int(gattr(ext, 'cy', 'a') or '0')
                                        else:
                                            pic_x = pic_y = 0
                                            pic_cx = cx
                                            pic_cy = cy
                                    else:
                                        pic_x = pic_y = 0
                                        pic_cx = cx
                                        pic_cy = cy

                                    abs_x = real_h + pic_x
                                    abs_y = real_v + pic_y

                                    x_mm = emu_to_mm(abs_x)
                                    y_mm = emu_to_mm(abs_y)
                                    w_mm = emu_to_mm(pic_cx)
                                    h_mm = emu_to_mm(pic_cy)

                                    z_idx = rel_height - 251640000
                                    html_str = (f'<div class="docx-element image" '
                                                f'style="position: absolute; left: {x_mm}mm; top: {y_mm}mm; '
                                                f'width: {w_mm}mm; height: {h_mm}mm; z-index: {z_idx};">'
                                                f'<img src="{img_src}" style="width:100%;height:100%;object-fit:fill;" />'
                                                f'</div>')
                                    all_elements.append((rel_height, behind_doc == '1', html_str))

            # 处理正文文字（非 anchor 的 run）
            p_html = parse_paragraph(child)
            if p_html:
                has_body_text = True
                y_mm = emu_to_mm(current_para_y)
                body_text_html = (f'<div class="docx-element body-text" '
                                  f'style="position: absolute; left: {emu_to_mm(mar_left_emu)}mm; top: {y_mm}mm; '
                                  f'width: {emu_to_mm(content_w_emu)}mm;">{p_html}</div>')
                body_content_elements.append((0, False, body_text_html))

            paragraph_y_emu += p_height

        elif tag == 'tbl':
            current_y = paragraph_y_emu
            tbl_html, tbl_height, tbl_floats = parse_table(
                child, current_y, mar_left_emu, page_w_emu, mar_top_emu, rels, image_data)
            if tbl_html:
                body_content_elements.append((0, False, tbl_html))
            body_content_elements.extend(tbl_floats)
            paragraph_y_emu += tbl_height

        elif tag == 'sectPr':
            pass

    # 合并所有元素
    all_elements.extend(body_content_elements)

    # 按 relativeHeight 排序
    all_elements.sort(key=lambda x: x[0])

    elements_html = '\n'.join(e[2] for e in all_elements)

    # 调试水印（仅 --debug 时显示，默认关闭，保证独立转换输出干净）
    page_info_div = (f'<div class="page-info">{page_w_mm:.1f}mm × {page_h_mm:.1f}mm | 页边距: {mar_left_mm:.1f}mm</div>'
                     if debug else '')

    # 生成 HTML
    html_content = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>DOCX 转 HTML</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            background: #e8e8e8;
            font-family: '微软雅黑', 'Microsoft YaHei', 'SimHei', 'SimSun', sans-serif;
            font-size: {default_font_pt:.2f}pt;
            {('line-height: %.3f;' % body_line_factor) if body_line_factor else ''}
            display: flex;
            justify-content: center;
            padding: 20px 0;
        }}

        .a4-page {{
            position: relative;
            width: {page_w_mm:.1f}mm;
            height: {page_h_mm:.1f}mm;
            background: white;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
            overflow: hidden;
        }}

        .docx-element {{
            position: absolute;
        }}

        .docx-element.text-box {{
            z-index: 200;
        }}
        .docx-element.label {{
            z-index: 100;
        }}
        .docx-element.body-text {{
            z-index: 200;
        }}
        .docx-element.table {{
            z-index: 200;
        }}
        .docx-element.image {{
            z-index: 50;
        }}
        .docx-element.shape {{
            z-index: 10;
        }}
        .docx-element.connector {{
            z-index: 15;
        }}
        .docx-element svg {{
            position: absolute;
            top: 0;
            left: 0;
            z-index: 0;
            pointer-events: none;
        }}
        .docx-element > div {{
            position: relative;
            z-index: 2;
        }}
        .docx-element > table {{
            position: relative;
            z-index: 2;
        }}

        .docx-element p {{
            margin: 0;
            padding: 0;
        }}

        .page-info {{
            position: fixed;
            top: 10px;
            right: 10px;
            background: rgba(0,0,0,0.7);
            color: white;
            padding: 8px 16px;
            border-radius: 6px;
            font-size: 12px;
            z-index: 9999;
        }}
    </style>
</head>
<body>
    {page_info_div}
    <div class="a4-page" id="a4Page">
{elements_html}
    </div>
</body>
</html>'''

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f'\n转换完成!')
    print(f'  输入: {docx_dir}')
    print(f'  输出: {output_path}')
    print(f'  元素数: {len(all_elements)}')
    print(f'  页面: {page_w_mm:.1f}mm × {page_h_mm:.1f}mm')
    print(f'  页边距: 上{mar_top_mm:.1f} 下{mar_bottom_mm:.1f} 左{mar_left_mm:.1f} 右{mar_right_mm:.1f}mm')

    return output_path


# ===================================================================
#  Web 上传界面
# ===================================================================
def start_web_server(port=8765):
    """启动一个简单的 Web 上传界面"""
    import http.server
    import socketserver

    UPLOAD_HTML = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>DOCX → HTML 转换器</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: '微软雅黑', sans-serif;
            background: #f0f2f5;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .container {
            background: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.1);
            width: 480px;
            text-align: center;
        }
        h1 { font-size: 22px; color: #333; margin-bottom: 8px; }
        .desc { color: #999; font-size: 14px; margin-bottom: 24px; }
        .upload-area {
            border: 2px dashed #ccc;
            border-radius: 8px;
            padding: 40px;
            cursor: pointer;
            transition: all 0.3s;
        }
        .upload-area:hover { border-color: #4F81BD; background: #f8faff; }
        .upload-area.dragover { border-color: #4F81BD; background: #eef5ff; }
        .upload-icon { font-size: 48px; color: #ccc; }
        .upload-text { color: #666; margin-top: 12px; }
        input[type=file] { display: none; }
        .btn {
            display: inline-block;
            padding: 12px 32px;
            background: #4F81BD;
            color: white;
            border: none;
            border-radius: 6px;
            font-size: 16px;
            cursor: pointer;
            margin-top: 20px;
            transition: background 0.3s;
        }
        .btn:hover { background: #3a6ba8; }
        .btn:disabled { background: #ccc; cursor: not-allowed; }
        .result { margin-top: 20px; display: none; }
        .result a { color: #4F81BD; text-decoration: none; font-size: 16px; }
        .result a:hover { text-decoration: underline; }
        .error { color: #d44; margin-top: 12px; display: none; }
        .progress { margin-top: 12px; color: #666; font-size: 14px; display: none; }
    </style>
</head>
<body>
    <div class="container">
        <h1>DOCX → HTML 转换器</h1>
        <p class="desc">上传 .docx 文件，自动 1:1 还原为 HTML</p>
        <div class="upload-area" id="dropZone" onclick="document.getElementById('fileInput').click()">
            <div class="upload-icon">&#128196;</div>
            <div class="upload-text">点击或拖拽 .docx 文件到此处</div>
        </div>
        <input type="file" id="fileInput" accept=".docx">
        <button class="btn" id="convertBtn" disabled onclick="convert()">开始转换</button>
        <div class="progress" id="progress">正在转换，请稍候...</div>
        <div class="result" id="result">
            <p>转换完成!</p>
            <a href="#" id="downloadLink" target="_blank">查看 HTML 结果</a>
        </div>
        <div class="error" id="error"></div>
    </div>
    <script>
        let selectedFile = null;
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const convertBtn = document.getElementById('convertBtn');
        const result = document.getElementById('result');
        const errorDiv = document.getElementById('error');
        const progress = document.getElementById('progress');

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                selectedFile = e.target.files[0];
                convertBtn.disabled = false;
                dropZone.querySelector('.upload-text').textContent = '已选择: ' + selectedFile.name;
            }
        });

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                selectedFile = e.dataTransfer.files[0];
                fileInput.files = e.dataTransfer.files;
                convertBtn.disabled = false;
                dropZone.querySelector('.upload-text').textContent = '已选择: ' + selectedFile.name;
            }
        });

        function convert() {
            if (!selectedFile) return;
            convertBtn.disabled = true;
            progress.style.display = 'block';
            result.style.display = 'none';
            errorDiv.style.display = 'none';

            const formData = new FormData();
            formData.append('file', selectedFile);

            fetch('/upload', { method: 'POST', body: formData })
                .then(r => r.json())
                .then(data => {
                    progress.style.display = 'none';
                    if (data.error) {
                        errorDiv.textContent = data.error;
                        errorDiv.style.display = 'block';
                    } else {
                        document.getElementById('downloadLink').href = '/output/' + data.filename;
                        result.style.display = 'block';
                    }
                    convertBtn.disabled = false;
                })
                .catch(err => {
                    progress.style.display = 'none';
                    errorDiv.textContent = '转换失败: ' + err.message;
                    errorDiv.style.display = 'block';
                    convertBtn.disabled = false;
                });
        }
    </script>
</body>
</html>'''

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            if self.path == '/' or self.path == '/index.html':
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(UPLOAD_HTML.encode('utf-8'))
            elif self.path.startswith('/output/'):
                filename = self.path.split('/output/')[-1]
                filepath = os.path.join(output_dir, filename)
                if os.path.exists(filepath):
                    self.send_response(200)
                    self.send_header('Content-Type', 'text/html; charset=utf-8')
                    self.end_headers()
                    with open(filepath, 'r', encoding='utf-8') as f:
                        self.wfile.write(f.read().encode('utf-8'))
                else:
                    self.send_error(404)
            else:
                self.send_error(404)

        def do_POST(self):
            if self.path == '/upload':
                content_type = self.headers.get('Content-Type', '')
                if 'multipart/form-data' not in content_type:
                    self._send_json({'error': '需要 multipart/form-data'})
                    return

                # 解析 multipart
                boundary = content_type.split('boundary=')[1].encode()
                content_length = int(self.headers.get('Content-Length', 0))
                body_data = self.rfile.read(content_length)

                # 提取文件
                parts = body_data.split(b'--' + boundary)
                file_data = None
                filename = 'upload.docx'

                for part in parts:
                    if b'filename=' in part:
                        # 提取文件名
                        header_end = part.find(b'\r\n\r\n')
                        if header_end > 0:
                            header = part[:header_end].decode('utf-8', errors='ignore')
                            if 'filename="' in header:
                                filename = header.split('filename="')[1].split('"')[0]
                            file_data = part[header_end + 4:]
                            # 去掉末尾的 \r\n
                            if file_data.endswith(b'\r\n'):
                                file_data = file_data[:-2]

                if file_data is None:
                    self._send_json({'error': '未找到文件'})
                    return

                # 保存上传文件
                upload_path = os.path.join(output_dir, filename)
                with open(upload_path, 'wb') as f:
                    f.write(file_data)

                # 转换
                output_name = os.path.splitext(filename)[0] + '.html'
                output_path = os.path.join(output_dir, output_name)

                try:
                    result = convert_docx_to_html(upload_path, output_path)
                    if result:
                        self._send_json({'success': True, 'filename': output_name})
                    else:
                        self._send_json({'error': '转换失败'})
                except Exception as e:
                    self._send_json({'error': str(e)})
            else:
                self.send_error(404)

        def _send_json(self, data):
            import json
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps(data).encode('utf-8'))

        def log_message(self, format, *args):
            pass  # 静默日志

    output_dir = os.path.join(os.getcwd(), 'output')
    os.makedirs(output_dir, exist_ok=True)

    with socketserver.TCPServer(("", port), Handler) as httpd:
        print(f'Web 服务器已启动: http://localhost:{port}')
        print(f'输出目录: {output_dir}')
        print(f'按 Ctrl+C 停止')
        httpd.serve_forever()


# ===================================================================
#  命令行入口
# ===================================================================
if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='DOCX → HTML 1:1 还原转换器',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例:
  python docx2html.py resume.docx              # 转换并输出到同目录
  python docx2html.py resume.docx -o out.html  # 指定输出路径
  python docx2html.py --web                     # 启动 Web 上传界面
  python docx2html.py --web --port 9000         # 指定端口
        '''
    )
    parser.add_argument('input', nargs='?', help='.docx 文件路径')
    parser.add_argument('-o', '--output', help='输出 HTML 文件路径')
    parser.add_argument('--web', action='store_true', help='启动 Web 上传界面')
    parser.add_argument('--port', type=int, default=8765, help='Web 服务端口 (默认 8765)')
    parser.add_argument('--debug', action='store_true', help='显示页面尺寸/页边距调试水印（默认关闭）')

    args = parser.parse_args()

    if args.web:
        start_web_server(args.port)
    elif args.input:
        convert_docx_to_html(args.input, args.output, debug=args.debug)
    else:
        parser.print_help()

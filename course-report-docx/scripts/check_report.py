#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""课程报告 / 毕业论文 .docx 格式自检

用法:
    python check_report.py 报告.docx

只依赖 Python 3 标准库（zipfile + xml.etree），不需要安装任何第三方包，
因此在任何有 Python 3 的环境里都能直接跑。

输出三类结果：
    PASS  符合规范
    WARN  非硬性要求，或无法判定（例如段落数不够取样）
    FAIL  违反硬性格式要求，必须修

退出码：0 = 无 FAIL；1 = 存在 FAIL。
"""

import sys
import os
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# ---------------------------------------------------------------- 规范常量
A4_W, A4_H = 11906, 16838                       # A4 纵向（缇）
BODY_MARGIN_TB, BODY_MARGIN_LR = 1418, 1701     # 正文节：上下 2.5cm / 左右 3.0cm
FRONT_MARGIN_TB, FRONT_MARGIN_LR = 1440, 1797   # 前置节：上下 2.54cm / 左右 3.17cm
HDR_DIST, FTR_DIST = 850, 992                   # 页眉 1.5cm / 页脚 1.75cm
TOL = 6                                         # 缇容差（1/100 mm 级别，避免换算取整误报）
SZ_BODY = 24                                    # 正文 小四 12pt（半磅）
SZ_CAPTION = 21                                 # 题注 五号 10.5pt
LINE_1_5 = 360                                  # 1.5 倍行距
FIRST_LINE_CHARS = 200                          # 首行缩进 2 字符
FIRST_LINE_TWIP = (420, 480)                    # 2 字符缩进的固定值写法（五号/小四）
TBL_BORDER_SZ = 4                               # 表格边框 4/8 磅 = 0.5 磅
HDR_BORDER_SZ = 6                               # 页眉下框线 6/8 磅 = 0.75 磅
EXPECTED_SECTIONS = 5                           # 封面 / 摘要 / 目录 / 正文 / 末节
CAPTION_STYLE = '5'                             # "题注"样式 styleId
HEADING_STYLES = {'2': '标题1', '3': '标题2', '4': '标题3'}
PLACEHOLDERS = ['XXXXXXX课程报告', '20XX级', 'XXXXXXXXXXXX', 'XXXXXX',
                '{{REPORT_TITLE}}', '{{DATE}}']
RED = {'ff0000'}
CAPTION_SEQS = {'SEQ 图': '图', 'SEQ 表': '表', 'SEQ 代码': '算法'}
MIN_BODY_LEN = 20                               # 取样正文段落的长度下限


class Report:
    def __init__(self):
        self.rows = []

    def _add(self, level, group, name, detail=''):
        self.rows.append((level, group, name, detail))

    def ok(self, g, n, d=''):
        self._add('PASS', g, n, d)

    def warn(self, g, n, d=''):
        self._add('WARN', g, n, d)

    def fail(self, g, n, d=''):
        self._add('FAIL', g, n, d)

    def render(self):
        icon = {'PASS': '[PASS]', 'WARN': '[WARN]', 'FAIL': '[FAIL]'}
        group = None
        for level, g, n, d in self.rows:
            if g != group:
                group = g
                print('\n== %s ==' % g)
            line = '  %s %s' % (icon[level], n)
            if d:
                line += '   ->  %s' % d
            print(line)
        c = {}
        for level, _, _, _ in self.rows:
            c[level] = c.get(level, 0) + 1
        fails = c.get('FAIL', 0)
        print('\n' + '-' * 64)
        print('合计 %d 项：PASS %d  WARN %d  FAIL %d'
              % (len(self.rows), c.get('PASS', 0), c.get('WARN', 0), fails))
        if fails:
            print('存在未通过的硬性格式要求，请先修正 FAIL 项再交付。')
        else:
            print('未发现硬性格式问题。')
        return fails


# ---------------------------------------------------------------- 取值辅助
def ival(el, attr='val', default=None):
    if el is None:
        return default
    v = el.get(W + attr) or el.get(attr)
    if v is None:
        return default
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def pr(p):
    return p.find(W + 'pPr')


def pstyle(p):
    pPr = pr(p)
    if pPr is None:
        return None
    s = pPr.find(W + 'pStyle')
    return s.get(W + 'val') if s is not None else None


def ptext(p):
    return ''.join(t.text or '' for t in p.iter(W + 't'))


def instr_text(p):
    return ' '.join(t.text or '' for t in p.iter(W + 'instrText'))


def has_tag(root, tag):
    for _ in root.iter(W + tag):
        return True
    return False


def body_blocks(body):
    """展开 body 的块级子元素，并钻进 w:sdt 内容控件。

    封面常被包在内容控件里，只看 body 的直接孩子会整段漏掉封面。
    """
    out = []
    for el in body:
        if el.tag == W + 'sdt':
            sc = el.find(W + 'sdtContent')
            if sc is not None:
                out.extend(body_blocks(sc))
            continue
        out.append(el)
    return out


def build_styles(z):
    """解析 styles.xml -> ({styleId: 属性}, 默认段落样式id, 文档默认字号)。支持 basedOn 继承。"""
    raw, default_id, doc_default_sz = {}, None, None
    if 'word/styles.xml' not in z.namelist():
        return {}, None, None
    root = ET.fromstring(z.read('word/styles.xml'))

    dd = root.find(W + 'docDefaults')
    if dd is not None:
        sz = dd.find(W + 'rPrDefault/' + W + 'rPr/' + W + 'sz')
        doc_default_sz = ival(sz)

    for st in root.iter(W + 'style'):
        if st.get(W + 'type') not in (None, 'paragraph'):
            continue
        sid = st.get(W + 'styleId')
        d = {'_basedOn': None}
        bo = st.find(W + 'basedOn')
        if bo is not None:
            d['_basedOn'] = bo.get(W + 'val')
        ppr, rpr = st.find(W + 'pPr'), st.find(W + 'rPr')
        if rpr is not None:
            sz = rpr.find(W + 'sz')
            if sz is not None:
                d['sz'] = ival(sz)
            b = rpr.find(W + 'b')
            if b is not None:
                d['b'] = b.get(W + 'val') not in ('0', 'false')
            rf = rpr.find(W + 'rFonts')
            if rf is not None:
                d['eastAsia'] = rf.get(W + 'eastAsia') or rf.get(W + 'eastAsiaTheme') or ''
        if ppr is not None:
            jc = ppr.find(W + 'jc')
            if jc is not None:
                d['jc'] = jc.get(W + 'val')
            sp = ppr.find(W + 'spacing')
            if sp is not None:
                if sp.get(W + 'line') is not None:
                    d['line'] = ival(sp, 'line')
                if sp.get(W + 'lineRule') is not None:
                    d['lineRule'] = sp.get(W + 'lineRule')
            ind = ppr.find(W + 'ind')
            if ind is not None:
                d['ind'] = {k.split('}')[-1]: v for k, v in ind.attrib.items()}
            if ppr.find(W + 'keepNext') is not None:
                d['keepNext'] = True
            if ppr.find(W + 'numPr') is not None:
                d['numPr'] = True
            pb = ppr.find(W + 'pBdr')
            if pb is not None:
                bo2 = pb.find(W + 'bottom')
                if bo2 is not None:
                    d['pBdrBottom'] = ival(bo2, 'sz')
        raw[sid] = d
        if st.get(W + 'default') in ('1', 'true'):
            default_id = sid

    def resolve(sid, depth=0):
        if sid is None or sid not in raw or depth > 8:
            return {}
        d = raw[sid]
        merged = dict(resolve(d.get('_basedOn'), depth + 1))
        for k, v in d.items():
            if k == '_basedOn':
                continue
            if k == 'ind' and isinstance(v, dict) and isinstance(merged.get('ind'), dict):
                m = dict(merged['ind'])
                m.update(v)
                merged['ind'] = m
            else:
                merged[k] = v
        return merged

    return {sid: resolve(sid) for sid in raw}, default_id, doc_default_sz


def para_style_d(p, styles, default_id):
    return styles.get(pstyle(p) or default_id, {})


def run_sizes(p, sd, default_sz):
    """段落中各可见 run 的有效字号集合（只返回能确定的整数）。"""
    out = set()
    for r in p.findall(W + 'r'):
        if not ''.join(t.text or '' for t in r.iter(W + 't')):
            continue
        rpr = r.find(W + 'rPr')
        sz = ival(rpr.find(W + 'sz')) if rpr is not None else None
        if sz is None:
            sz = sd.get('sz', default_sz)
        if sz is not None:
            out.add(sz)
    return out


def run_bolds(p, sd):
    out = []
    for r in p.findall(W + 'r'):
        if not ''.join(t.text or '' for t in r.iter(W + 't')):
            continue
        rpr = r.find(W + 'rPr')
        b = rpr.find(W + 'b') if rpr is not None else None
        out.append(b.get(W + 'val') not in ('0', 'false') if b is not None
                   else bool(sd.get('b')))
    return out


def eff_jc(p, sd):
    pPr = pr(p)
    if pPr is not None:
        jc = pPr.find(W + 'jc')
        if jc is not None:
            return jc.get(W + 'val')
    return sd.get('jc')


def eff_line(p, sd):
    pPr = pr(p)
    if pPr is not None:
        sp = pPr.find(W + 'spacing')
        if sp is not None and sp.get(W + 'line') is not None:
            return ival(sp, 'line'), sp.get(W + 'lineRule')
    return sd.get('line'), sd.get('lineRule')


def eff_ind(p, sd):
    pPr = pr(p)
    if pPr is not None:
        ind = pPr.find(W + 'ind')
        if ind is not None:
            return {k.split('}')[-1]: v for k, v in ind.attrib.items()}
    return sd.get('ind', {}) or {}


def eff_keepnext(p, sd):
    pPr = pr(p)
    if pPr is not None and pPr.find(W + 'keepNext') is not None:
        return True
    return bool(sd.get('keepNext'))


def to_int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def indent_ok(ind):
    """首行缩进 2 字符：firstLineChars=200，或固定值 420/480 缇。"""
    flc = to_int(ind.get('firstLineChars'))
    if flc == FIRST_LINE_CHARS:
        return True
    return to_int(ind.get('firstLine')) in FIRST_LINE_TWIP


def is_caption(p):
    """只有含 SEQ 域的段落才算真题注；定义/定理段也用"题注"样式但不是题注。"""
    ins = instr_text(p)
    return any(seq in ins for seq in CAPTION_SEQS)


def caption_kind(p):
    ins = instr_text(p)
    for seq, label in CAPTION_SEQS.items():
        if seq in ins:
            return label
    return None


def is_figure_container(p):
    """图文框/文本框容器段落：文字在文本框里，不按正文规则检查。"""
    if p.find('.//' + W + 'txbxContent') is not None:
        return True
    return (p.find('.//' + W + 'drawing') is not None
            or p.find('.//' + W + 'pict') is not None)


# ---------------------------------------------------------------- 主检查
def check(path):
    r = Report()

    if not os.path.isfile(path):
        r.fail('文件', '目标文件存在', path)
        return r
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        r.fail('文件', '是合法的 .docx（zip）', '无法解压，文件可能损坏或加密')
        return r

    names = z.namelist()
    need = ['word/document.xml', 'word/styles.xml', 'word/numbering.xml',
            'word/settings.xml']
    missing = [n for n in need if n not in names]
    if missing:
        r.fail('文件', 'OOXML 必需部件齐全', '缺少 ' + ', '.join(missing))
        return r
    r.ok('文件', 'OOXML 必需部件齐全', '%d 个部件' % len(names))

    doc = ET.fromstring(z.read('word/document.xml'))
    styles, default_id, default_sz = build_styles(z)
    body = doc.find(W + 'body')
    blocks = body_blocks(body)
    direct_ps = [el for el in blocks if el.tag == W + 'p']

    # 封面区 = 第一个"承载分节符"的段落及其之前的所有段落。
    # 封面用的是 62pt 大标题 + 16pt 表格文字，绝不能拿去比对正文格式。
    cover_ps, has_break = [], False
    for p in direct_ps:
        cover_ps.append(p)
        if p.find(W + 'pPr/' + W + 'sectPr') is not None:
            has_break = True
            break
    if not has_break:                 # 没有分节符，界线不明，不做排除
        cover_ps = []
    cover_ids = {id(p) for p in cover_ps}

    # ------------------------------------------------ 页面与分节
    sects = list(doc.iter(W + 'sectPr'))
    n = len(sects)
    if n == EXPECTED_SECTIONS:
        r.ok('页面与分节', '分节数为 %d（封面/摘要/目录/正文/末节）' % n)
    else:
        r.warn('页面与分节', '分节数为 %d，规范期望 %d' % (n, EXPECTED_SECTIONS),
               '分节不对会导致页码与页眉失控')

    pgsz = None
    for sp in sects:
        el = sp.find(W + 'pgSz')
        if el is not None and ival(el, 'w') and ival(el, 'h'):
            pgsz = (ival(el, 'w'), ival(el, 'h'))
    if pgsz == (A4_W, A4_H):
        r.ok('页面与分节', '纸张 A4 纵向', '%d x %d 缇' % pgsz)
    elif pgsz:
        r.fail('页面与分节', '纸张 A4 纵向', '实测 %d x %d 缇' % pgsz)
    else:
        r.warn('页面与分节', '纸张 A4 纵向', '未读到 pgSz')

    def near(a, b):
        return a is not None and abs(a - b) <= TOL

    front_bad, body_ok, dist_bad = 0, False, 0
    for i, sp in enumerate(sects):
        m = sp.find(W + 'pgMar')
        if m is None:
            continue
        t, b = ival(m, 'top'), ival(m, 'bottom')
        l, ri = ival(m, 'left'), ival(m, 'right')
        if not near(ival(m, 'header'), HDR_DIST) or not near(ival(m, 'footer'), FTR_DIST):
            dist_bad += 1
        if i == n - 1:
            continue                      # 末节为空节，不参与版式比对
        if i == n - 2:                    # 正文节
            body_ok = (near(t, BODY_MARGIN_TB) and near(b, BODY_MARGIN_TB)
                       and near(l, BODY_MARGIN_LR) and near(ri, BODY_MARGIN_LR))
        else:                             # 封面 / 摘要 / 目录
            if not (near(t, FRONT_MARGIN_TB) and near(b, FRONT_MARGIN_TB)
                    and near(l, FRONT_MARGIN_LR) and near(ri, FRONT_MARGIN_LR)):
                front_bad += 1

    if body_ok:
        r.ok('页面与分节', '正文节页边距 2.5cm / 3.0cm')
    else:
        r.fail('页面与分节', '正文节页边距 2.5cm / 3.0cm', '应为 1418 / 1701 缇')
    if front_bad == 0:
        r.ok('页面与分节', '前置节页边距 2.54cm / 3.17cm')
    else:
        r.warn('页面与分节', '前置节页边距 2.54cm / 3.17cm', '%d 个节不符' % front_bad)
    if dist_bad == 0:
        r.ok('页面与分节', '页眉/页脚距边界 1.5cm / 1.75cm')
    else:
        r.warn('页面与分节', '页眉/页脚距边界 1.5cm / 1.75cm', '%d 个节不符' % dist_bad)

    # ------------------------------------------------ 页码与页眉
    fmt_start = []
    for sp in sects:
        p = sp.find(W + 'pgNumType')
        fmt_start.append((p.get(W + 'fmt') if p is not None else None,
                          ival(p, 'start') if p is not None else None))
    if any(f == 'upperRoman' for f, _ in fmt_start):
        r.ok('页码与页眉', '目录节页码为罗马数字')
    else:
        r.warn('页码与页眉', '目录节页码为罗马数字', '未找到 fmt=upperRoman')
    if any(s == 1 for _, s in fmt_start):
        r.ok('页码与页眉', '正文节页码从 1 开始')
    else:
        r.warn('页码与页眉', '正文节页码从 1 开始', '未找到 start=1')

    # 封面节（第一节）不得带页脚引用与页码设置，否则封面会印出"I"
    if not sects:
        r.warn('页码与页眉', '封面不显示页码', '未读到节属性，无法判定')
    else:
        c0 = sects[0]
        bad = []
        if c0.find(W + 'footerReference') is not None:
            bad.append('封面节带 footerReference（页脚里有 PAGE 域）')
        if c0.find(W + 'headerReference') is not None:
            bad.append('封面节带 headerReference')
        if c0.find(W + 'pgNumType') is not None:
            f = c0.find(W + 'pgNumType')
            bad.append('封面节设了 pgNumType fmt=%s start=%s'
                       % (f.get(W + 'fmt') or '-', ival(f, 'start')))
        page_in_cover = [p for p in cover_ps
                         if 'PAGE' in instr_text(p).replace('NUMPAGES', '')]
        if page_in_cover:
            bad.append('封面段落内嵌 PAGE 域')
        if bad:
            r.fail('页码与页眉', '封面不显示页码', '；'.join(bad))
        else:
            r.ok('页码与页眉', '封面不显示页码', '封面节无页脚引用、无 pgNumType')

    hdr_parts = [x for x in names if x.startswith('word/header')]
    hdr_ref = sum(1 for sp in sects if sp.find(W + 'headerReference') is not None)
    if hdr_parts and hdr_ref:
        r.ok('页码与页眉', '存在页眉部件与页眉引用',
             '%d 个页眉部件 / %d 处引用' % (len(hdr_parts), hdr_ref))
    else:
        r.warn('页码与页眉', '存在页眉部件与页眉引用', '未检出')

    ph_hit, hdr_seen, hdr_ok_n = [], 0, 0
    for hp in hdr_parts:
        try:
            hr = ET.fromstring(z.read(hp))
        except ET.ParseError:
            continue
        for p in hr.iter(W + 'p'):
            t = ptext(p).strip()
            if not t:
                continue
            for ph in PLACEHOLDERS:
                if ph in t:
                    ph_hit.append('%s：%s' % (hp.split('/')[-1], ph))
            hdr_seen += 1
            sd = para_style_d(p, styles, default_id)
            jc, bot = eff_jc(p, sd), None
            pPr = pr(p)
            if pPr is not None:
                pb = pPr.find(W + 'pBdr')
                if pb is not None:
                    bot = ival(pb.find(W + 'bottom'), 'sz')
            if bot is None:
                bot = sd.get('pBdrBottom')
            if jc == 'center' and bot == HDR_BORDER_SZ:
                hdr_ok_n += 1

    if ph_hit:
        r.fail('页码与页眉', '页眉占位符已替换为真实课程名', '；'.join(ph_hit))
    else:
        r.ok('页码与页眉', '页眉占位符已替换为真实课程名')
    if hdr_seen and hdr_ok_n == hdr_seen:
        r.ok('页码与页眉', '页眉居中且带 0.75 磅下框线', '%d 处' % hdr_ok_n)
    elif hdr_seen:
        r.warn('页码与页眉', '页眉居中且带 0.75 磅下框线',
               '%d/%d 处符合' % (hdr_ok_n, hdr_seen))
    else:
        r.warn('页码与页眉', '页眉居中且带 0.75 磅下框线', '页眉中无文字，跳过')

    # ------------------------------------------------ 正文格式
    para_ps = []
    for p in direct_ps:
        if id(p) in cover_ids:          # 封面：62pt 大标题、16pt 表格文字，不按正文算
            continue
        s = pstyle(p)
        if s in HEADING_STYLES or s == CAPTION_STYLE:
            continue
        if len(ptext(p).strip()) < MIN_BODY_LEN:
            continue
        if is_figure_container(p):
            continue
        if any(seq in instr_text(p) for seq in CAPTION_SEQS):
            continue
        para_ps.append(p)

    if not para_ps:
        r.warn('正文格式', '正文段落取样', '未找到足够长的正文段落')
    else:
        sd_of = lambda p: para_style_d(p, styles, default_id)

        bad_sz = [p for p in para_ps
                  if (run_sizes(p, sd_of(p), default_sz) or {SZ_BODY}) != {SZ_BODY}]
        if not bad_sz:
            r.ok('正文格式', '正文字号小四 12pt', '%d 段全部符合' % len(para_ps))
        else:
            r.fail('正文格式', '正文字号小四 12pt',
                   '%d/%d 段不符（如：%s…）'
                   % (len(bad_sz), len(para_ps), ptext(bad_sz[0]).strip()[:24]))

        bad_line = [p for p in para_ps if eff_line(p, sd_of(p)) != (LINE_1_5, 'auto')]
        if not bad_line:
            r.ok('正文格式', '正文行距 1.5 倍', '%d 段' % len(para_ps))
        else:
            r.warn('正文格式', '正文行距 1.5 倍',
                   '%d/%d 段不符' % (len(bad_line), len(para_ps)))

        bad_jc = [p for p in para_ps if eff_jc(p, sd_of(p)) not in (None, 'both')]
        if not bad_jc:
            r.ok('正文格式', '正文两端对齐', '%d 段' % len(para_ps))
        else:
            r.fail('正文格式', '正文两端对齐',
                   '%d/%d 段不符' % (len(bad_jc), len(para_ps)))

        bad_ind = [p for p in para_ps if not indent_ok(eff_ind(p, sd_of(p)))]
        if not bad_ind:
            r.ok('正文格式', '正文首行缩进 2 字符', '%d 段' % len(para_ps))
        else:
            r.warn('正文格式', '正文首行缩进 2 字符',
                   '%d/%d 段不符（模板中存在 420 缇与 2 字符两种写法，均按 2 字符放行；'
                   '此处为完全无缩进的段落）' % (len(bad_ind), len(para_ps)))

    # ------------------------------------------------ 标题
    found = {s: 0 for s in HEADING_STYLES}
    numpr = 0
    for p in direct_ps:
        s = pstyle(p)
        if s in HEADING_STYLES:
            found[s] += 1
            pPr = pr(p)
            if pPr is not None and pPr.find(W + 'numPr') is not None:
                numpr += 1
            elif styles.get(s, {}).get('numPr'):
                numpr += 1
    total_h = sum(found.values())
    if total_h:
        r.ok('标题', '使用内置标题 1/2/3 样式',
             ' / '.join('%s=%d' % (v, found[k]) for k, v in HEADING_STYLES.items()))
        if numpr == total_h:
            r.ok('标题', '标题挂接多级列表自动编号', '%d 处' % numpr)
        else:
            r.fail('标题', '标题挂接多级列表自动编号',
                   '仅 %d/%d 处带 numPr' % (numpr, total_h))
    else:
        r.warn('标题', '使用内置标题 1/2/3 样式', '未检出标题段落')

    # ------------------------------------------------ 题注
    caps = [p for p in direct_ps if pstyle(p) == CAPTION_STYLE and is_caption(p)]
    if not caps:
        r.warn('题注', '题注格式统一', '未检出带 SEQ 域的题注段落')
    else:
        bad_align, bad_bold, bad_sz, bad_ind, bad_kn = [], [], [], [], []
        kinds = {'图': 0, '表': 0, '算法': 0}
        for p in caps:
            sd = para_style_d(p, styles, default_id)
            k = caption_kind(p)
            if k in kinds:
                kinds[k] += 1
            if eff_jc(p, sd) != 'center':
                bad_align.append(p)
            szs = run_sizes(p, sd, default_sz)
            if szs and szs != {SZ_CAPTION}:
                bad_sz.append(p)
            bs = run_bolds(p, sd)
            if bs and not all(bs):
                bad_bold.append(p)
            ind = eff_ind(p, sd)
            if to_int(ind.get('firstLine')) or to_int(ind.get('firstLineChars')):
                bad_ind.append(p)
            if k == '图' and eff_keepnext(p, sd):
                bad_kn.append(p)
            if k in ('表', '算法') and not eff_keepnext(p, sd):
                bad_kn.append(p)

        total = len(caps)
        r.ok('题注', '识别到题注', '图 %d / 表 %d / 算法 %d'
             % (kinds['图'], kinds['表'], kinds['算法']))

        def verdict(bad, name, msg):
            if not bad:
                r.ok('题注', name, '%d 处全部符合' % total)
            else:
                r.fail('题注', name, '%d/%d 处%s' % (len(bad), total, msg))

        verdict(bad_align, '题注居中', '不居中')
        verdict(bad_bold, '题注加粗（图/表/算法统一）', '未加粗')
        verdict(bad_sz, '题注字号五号 10.5pt', '字号不符')
        verdict(bad_ind, '题注不首行缩进', '仍带缩进')
        verdict(bad_kn, '"与下段同页"（图题注不设，表/算法题注设）', '设置不符')

    # 图片段落 keepNext：只检查"紧邻图题注"的图片段落
    img_figs = 0
    img_kn_bad = 0
    for i, p in enumerate(direct_ps):
        if not is_figure_container(p):
            continue
        nxt = None
        for j in range(i + 1, min(i + 3, len(direct_ps))):
            if ptext(direct_ps[j]).strip():
                nxt = direct_ps[j]
                break
        if nxt is None or caption_kind(nxt) != '图':
            continue
        img_figs += 1
        if not eff_keepnext(p, para_style_d(p, styles, default_id)):
            img_kn_bad += 1
    if img_figs:
        if img_kn_bad == 0:
            r.ok('题注', '图片段落设"与下段同页"', '%d 处' % img_figs)
        else:
            r.fail('题注', '图片段落设"与下段同页"',
                   '%d/%d 处缺失，图和题注可能被分页断开' % (img_kn_bad, img_figs))
    else:
        r.warn('题注', '图片段落设"与下段同页"', '未找到紧邻图题注的图片段落')

    # ------------------------------------------------ 表格与域
    # 封面信息表是无框线的浮动表格（tblpPr 页级锚定），不参与全框线检查
    cover_tbls = set()
    if has_break:
        for el in blocks:
            if el.tag == W + 'tbl':
                cover_tbls.add(id(el))
            elif el.tag == W + 'p' and el.find(W + 'pPr/' + W + 'sectPr') is not None:
                break
    tbls = [t for t in body.iter(W + 'tbl') if id(t) not in cover_tbls]
    if tbls:
        bad_t = 0
        for t in tbls:
            borders = t.find(W + 'tblPr/' + W + 'tblBorders')
            if borders is None:
                bad_t += 1
                continue
            if not any(ival(borders.find(W + s), 'sz') == TBL_BORDER_SZ
                       for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')):
                bad_t += 1
        if bad_t == 0:
            r.ok('表格与域', '表格为全框线 0.5 磅', '%d 张表（封面信息表已排除）' % len(tbls))
        else:
            r.warn('表格与域', '表格为全框线 0.5 磅',
                   '%d/%d 张不符（封面信息表已排除）' % (bad_t, len(tbls)))
    else:
        r.warn('表格与域', '表格为全框线 0.5 磅', '正文中没有表格')

    doc_instr = ' '.join(instr_text(p) for p in direct_ps)
    if 'TOC' in doc_instr:
        r.ok('表格与域', '目录由 TOC 域自动生成')
    else:
        r.fail('表格与域', '目录由 TOC 域自动生成', '未找到 TOC 域，目录可能是手打的')
    for seq, label in (('SEQ 图', '图'), ('SEQ 表', '表')):
        if seq in doc_instr:
            r.ok('表格与域', '%s题注使用自动编号域' % label)
        else:
            r.warn('表格与域', '%s题注使用自动编号域' % label, '未检出')

    # ------------------------------------------------ 结构完整性
    # 这些问题 Word 打不开或报"文件可能已经损坏"，必须硬性拦截
    empty_tc = [tc for tc in doc.iter(W + 'tc') if tc.find(W + 'p') is None]
    if empty_tc:
        r.fail('结构完整性', '每个单元格至少含一个段落',
               '%d 个 w:tc 为空，Word 会报"文件可能已经损坏"（多为误删了单元格里最后一个空段落）'
               % len(empty_tc))
    else:
        r.ok('结构完整性', '每个单元格至少含一个段落')

    last = list(body)[-1] if len(body) else None
    if last is not None and last.tag == W + 'sectPr':
        r.ok('结构完整性', 'body 以 sectPr 结尾')
    else:
        r.fail('结构完整性', 'body 以 sectPr 结尾',
               'body 最后一个元素是 %s，文档缺少末节属性'
               % (last.tag.split('}')[-1] if last is not None else '空'))

    bad_tbl2 = [t for t in doc.iter(W + 'tbl') if t.find(W + 'tr') is None]
    if bad_tbl2:
        r.fail('结构完整性', '表格至少含一行', '%d 张空表' % len(bad_tbl2))
    else:
        r.ok('结构完整性', '表格至少含一行')

    anchors = sum(1 for t in ('commentRangeStart', 'commentRangeEnd',
                              'commentReference') for _ in doc.iter(W + t))
    has_comments = 'word/comments.xml' in names
    if anchors and not has_comments:
        r.fail('结构完整性', '批注部件与正文锚点一致',
               'comments.xml 已删除，正文却还剩 %d 个批注锚点；Word 打开会报'
               '"文件可能已经损坏"' % anchors)
    elif anchors:
        r.ok('结构完整性', '批注部件与正文锚点一致',
             '%d 个锚点，comments.xml 在（尚未清批注）' % anchors)
    else:
        r.ok('结构完整性', '批注部件与正文锚点一致')

    # ------------------------------------------------ 残留清理
    red_runs = 0
    for p in direct_ps:
        for r_el in p.findall(W + 'r'):
            if not ''.join(t.text or '' for t in r_el.iter(W + 't')).strip():
                continue
            c = r_el.find(W + 'rPr/' + W + 'color')
            if c is not None and (c.get(W + 'val') or '').lower() in RED:
                red_runs += 1
    if red_runs == 0:
        r.ok('残留清理', '没有红色说明文字残留')
    else:
        r.fail('残留清理', '没有红色说明文字残留',
               '%d 处红色 run 未删除（模板作者的排版说明）' % red_runs)

    hit = sorted({ph for ph in PLACEHOLDERS
                  if any(ph in ptext(p) for p in direct_ps)})
    if hit:
        r.fail('残留清理', '封面占位符已替换', '仍存在 ' + ', '.join(hit))
    else:
        r.ok('残留清理', '封面占位符已替换')

    if 'word/comments.xml' in names:
        r.fail('残留清理', '没有批注残留', '存在 word/comments.xml，请删除所有批注')
    else:
        r.ok('残留清理', '没有批注残留')

    if has_tag(doc, 'ins') or has_tag(doc, 'del'):
        r.fail('残留清理', '没有修订痕迹残留', '存在 w:ins / w:del，请接受所有修订')
    else:
        r.ok('残留清理', '没有修订痕迹残留')

    return r


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    path = sys.argv[1]
    print('检查文件：%s' % path)
    fails = check(path).render()
    sys.exit(1 if fails else 0)


if __name__ == '__main__':
    main()

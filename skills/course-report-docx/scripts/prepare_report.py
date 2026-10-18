#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""课程报告模板预处理：把模板变成一份可直接填内容的"干净起点"

用法:
    python prepare_report.py 源模板.docx 输出.docx --course "数据库课程报告" \
        --title "基于 REA 的 AI 辅助逆向分析实验开题报告" --date "2026年10月"

    # 只清模板、不填封面（保留 {{REPORT_TITLE}} / {{DATE}} 占位符）
    python prepare_report.py 源模板.docx 输出.docx --course "数据库课程报告"

它会做这些机械化的事（这些步骤靠人工在 Word 里做又慢又容易漏）：
    1. 删除所有红色（FF0000）说明文字，并清掉因此变空的段落
    2. 删除批注（comments.xml 及其关系、Content_Types 条目）
    3. 接受所有修订（展开 w:ins，丢弃 w:del）
    4. 把页眉里的 XXXXXXX课程报告 替换为 --course 指定的课程名
    5. 把封面的 {{REPORT_TITLE}} / {{DATE}} 替换为 --title / --date
    6. 统一题注格式：图/表/算法题注一律居中、加粗、不缩进；
       表与算法题注设"与下段同页"，图题注不设
    7. 给紧邻图题注的图片段落补上"与下段同页"
    8. 给无样式的正文段落补齐小四 12pt（模板的 Normal 样式是五号，
       不补的话直接打字会得到五号字）

它**不会**做的事（必须人工完成，因为没有可靠的判定依据）：
    - 删除示例插图、示例表格、定义/定理示例、算法示例框
    - 改写封面里的姓名、学号、学院、专业、年级
    - 撰写正文内容

源文件永远只读，输出写到新路径。只依赖 Python 3 标准库。
"""

import sys
import re
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
PR = '{http://schemas.openxmlformats.org/package/2006/relationships}'

RED = {'ff0000'}
SZ_BODY = 24
SZ_CAPTION = 21
PH_TITLE = '{{REPORT_TITLE}}'
PH_DATE = '{{DATE}}'
HDR_PLACEHOLDER = 'XXXXXXX课程报告'
CAPTION_SEQS = {'SEQ 图': '图', 'SEQ 表': '表', 'SEQ 代码': '算法'}
MIN_BODY_LEN = 20

# CT_PPr / CT_RPr 的子元素顺序（ECMA-376），插错位置 Word 会报"内容有问题"
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
             'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
             'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
             'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
             'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
             'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
             'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl',
             'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']
RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
             'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
             'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
             'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
             'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
             'eastAsianLayout', 'specVanish', 'oMath',
             'ins', 'del', 'moveFrom', 'moveTo', 'rPrChange']

REMOVE_TAGS = ['pPrChange', 'rPrChange', 'tblPrChange', 'trPrChange',
               'tcPrChange', 'sectPrChange', 'del', 'moveFrom', 'cellDel',
               'tblGridChange', 'numberingChange']
UNWRAP_TAGS = ['ins', 'moveTo', 'cellIns']
COMMENT_TAGS = ['commentRangeStart', 'commentRangeEnd', 'commentReference']


# ---------------------------------------------------------------- 基础工具
def q(tag):
    return W + tag


def register_ns(raw):
    """把原文件里声明的命名空间前缀原样注册，避免序列化时丢前缀。"""
    head = raw[:8000].decode('utf-8', 'ignore')
    for pfx, uri in re.findall(r'xmlns:([A-Za-z0-9_.-]+)="([^"]+)"', head):
        if pfx in ('xml', 'xmlns'):
            continue
        try:
            ET.register_namespace(pfx, uri)
        except (ValueError, TypeError):
            pass


def load(raw):
    register_ns(raw)
    return ET.fromstring(raw)


def root_open_tag(s):
    i = 0
    while True:
        i = s.find('<', i)
        if i < 0:
            return None
        if len(s) > i + 1 and s[i + 1] not in '?!':
            break
        i += 1
    return s[i:s.find('>', i) + 1]


def serialize(root, orig_raw=None):
    """序列化，并补回被 ElementTree 裁掉的命名空间声明。

    ET 只输出"实际用到"的命名空间，但源文件里 mc:Ignorable="w14 w15 wp14"
    这类属性会引用未使用的前缀；前缀未声明时严格解析器可能报错。
    这里把原文件根元素上声明的所有前缀原样补回去。
    """
    s = ET.tostring(root, encoding='UTF-8', xml_declaration=True).decode('utf-8')
    s = re.sub(r'^<\?xml[^>]*\?>',
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', s)

    if orig_raw is not None:
        decls = re.findall(r'xmlns:([A-Za-z0-9_.-]+)="([^"]+)"',
                           orig_raw[:8000].decode('utf-8', 'ignore'))
        tag = root_open_tag(s)
        if tag:
            add = ''.join(' xmlns:%s="%s"' % (p, u) for p, u in decls
                          if ('xmlns:%s="' % p) not in tag)
            if add:
                fixed = tag[:-1].rstrip() + add + ('/>' if tag.endswith('/>') else '>')
                s = s.replace(tag, fixed, 1)
    return s.encode('utf-8')


def parent_map(root):
    return {c: p for p in root.iter() for c in p}


def pPr_has_table_ancestor(el, pm):
    """判断元素是否位于表格单元格内。"""
    node = pm.get(el)
    while node is not None:
        if node.tag == q('tbl'):
            return True
        node = pm.get(node)
    return False


def localname(tag):
    return tag.split('}')[-1]


def ordered_insert(parent, child, order):
    """按 OOXML 规定的子元素顺序插入（顺序错了 Word 会报'内容有问题'）。"""
    lt = localname(child.tag)
    idx = order.index(lt) if lt in order else len(order)
    for i, existing in enumerate(parent):
        le = localname(existing.tag)
        eidx = order.index(le) if le in order else len(order)
        if eidx > idx:
            parent.insert(i, child)
            return child
    parent.append(child)
    return child


def set_child(parent, tag, attrs, order):
    el = parent.find(tag)
    if el is None:
        el = ordered_insert(parent, ET.Element(tag), order)
    for k, v in attrs.items():
        el.set(W + k, v)
    return el


def drop_child(parent, tag):
    el = parent.find(tag)
    if el is not None:
        parent.remove(el)
    return el


def ensure_pPr(p):
    pPr = p.find(q('pPr'))
    if pPr is None:
        pPr = ET.Element(q('pPr'))
        p.insert(0, pPr)
    return pPr


def ensure_rPr(r):
    rPr = r.find(q('rPr'))
    if rPr is None:
        rPr = ET.Element(q('rPr'))
        r.insert(0, rPr)
    return rPr


def ptext(p):
    return ''.join(t.text or '' for t in p.iter(q('t')))


def instr_text(p):
    return ' '.join(t.text or '' for t in p.iter(q('instrText')))


def visible_runs(el):
    out = []
    for r in el.iter(q('r')):
        if ''.join(t.text or '' for t in r.iter(q('t'))):
            out.append(r)
    return out


def is_red(run):
    c = run.find(q('rPr') + '/' + q('color'))
    return c is not None and (c.get(W + 'val') or '').lower() in RED


def caption_kind(p):
    ins = instr_text(p)
    for seq, label in CAPTION_SEQS.items():
        if seq in ins:
            return label
    return None


def has_figure(p):
    return (p.find('.//' + q('drawing')) is not None
            or p.find('.//' + q('pict')) is not None
            or p.find('.//' + q('txbxContent')) is not None)


# ---------------------------------------------------------------- 各个处理步骤
def is_last_para_in_cell(p, pm):
    """p 是否是所在单元格 / 文本框里唯一剩下的段落。

    w:tc 必须至少含一个 w:p，且必须以 w:p 结尾；把单元格里最后一个空段落删掉，
    Word 会直接报"文件可能已经损坏"。这是本脚本最容易踩的坑。
    """
    node = pm.get(p)
    while node is not None and node.tag not in (q('tc'), q('txbxContent')):
        node = pm.get(node)
    if node is None:
        return False
    return len(list(node.iter(q('p')))) <= 1


def strip_red_runs(root, stats):
    """删除红色说明文字，并清掉"因它才变空"的段落。

    只清理被红字牵连变空的段落。模板里本来就有意留空的段落（封面用 22 个空段落
    撑版面）必须原样保留，否则封面会挤成一团、日期顶到信息表上方。
    """
    touched = set()
    for _ in range(3):
        pm = parent_map(root)
        changed = False
        for r in [e for e in root.iter(q('r'))]:
            if is_red(r) and r in pm:
                node = pm[r]
                while node is not None and node.tag != q('p'):
                    node = pm.get(node)
                if node is not None:
                    touched.add(node)
                pm[r].remove(r)
                stats['red_runs'] += 1
                changed = True
        if not changed:
            break

    # 清空段落：只处理上面记录到的段落
    pm = parent_map(root)
    for p in list(touched):
        if p not in pm:
            continue
        if ptext(p).strip():
            continue
        if has_figure(p):
            continue
        pPr = p.find(q('pPr'))
        if pPr is not None and pPr.find(q('sectPr')) is not None:
            continue                      # 承载分节符的段落不能删
        if p.find(q('r')) is not None or p.find(q('hyperlink')) is not None:
            continue                      # 还有非文字内容（如空 run 带格式）
        if is_last_para_in_cell(p, pm):
            continue                      # 单元格里仅剩的段落必须保留
        pm[p].remove(p)
        stats['empty_paras'] += 1


def strip_comment_markers(root, stats):
    """删除正文里的批注锚点（commentRangeStart / End / commentReference）。

    模板里的批注靠 comments.xml 承载，正文只放锚点。一旦把 comments.xml
    删掉而正文锚点还在，Word 打开时会直接报"文件可能已经损坏"。
    """
    tags = {q(t) for t in COMMENT_TAGS}
    pm = parent_map(root)
    hit = [e for e in root.iter() if e.tag in tags and e in pm]
    dirty = []
    for e in hit:
        par = pm[e]
        par.remove(e)
        stats['comment_marks'] = stats.get('comment_marks', 0) + 1
        if par.tag == q('r'):
            dirty.append(par)

    # 锚点被摘掉后只剩 rPr 的 run，整条一起删，避免留下空壳
    if dirty:
        pm2 = parent_map(root)
        for r in dirty:
            if r not in pm2:
                continue
            if not [c for c in r if c.tag != q('rPr')]:
                pm2[r].remove(r)


def accept_revisions(root, stats):
    """展开 w:ins / w:moveTo，删除 w:del / 各类 *Change。"""
    for _ in range(4):
        pm = parent_map(root)
        changed = False
        for e in list(root.iter()):
            if e not in pm:
                continue
            local = e.tag.split('}')[-1]
            parent = pm[e]
            if local in REMOVE_TAGS:
                parent.remove(e)
                stats['rev_removed'] += 1
                changed = True
            elif local in UNWRAP_TAGS:
                idx = list(parent).index(e)
                for child in list(e):
                    e.remove(child)
                    parent.insert(idx, child)
                    idx += 1
                parent.remove(e)
                stats['rev_unwrapped'] += 1
                changed = True
        if not changed:
            break


def replace_across_runs(p, old, new, limit=8):
    """在段落内跨 run 替换文字（Word 常把一个词拆到多个 w:t 里）。"""
    cnt = 0
    for _ in range(limit):
        ts = list(p.iter(q('t')))
        if not ts:
            break
        spans, full, pos = [], '', 0
        for t in ts:
            s = t.text or ''
            spans.append((t, pos, pos + len(s)))
            full += s
            pos += len(s)
        i = full.find(old)
        if i < 0:
            break
        j = i + len(old)
        placed = False
        for t, a, b in spans:
            if b <= i or a >= j:
                continue
            seg = t.text or ''
            ls, le = max(a, i) - a, min(b, j) - a
            t.text = seg[:ls] + (new if not placed else '') + seg[le:]
            placed = True
        cnt += 1
    return cnt


def replace_header_placeholder(root, course, stats):
    for p in root.iter(q('p')):
        stats['header_replaced'] += replace_across_runs(p, HDR_PLACEHOLDER, course)


def fill_placeholders(root, title, date, stats):
    """替换封面的 {{REPORT_TITLE}} / {{DATE}}（跨 run 安全）。

    返回 (找到的占位符数, 成功替换数)，供调用方判断是否还有残留。
    """
    for p in root.iter(q('p')):
        if title:
            stats['title_replaced'] += replace_across_runs(p, PH_TITLE, title)
        if date:
            stats['date_replaced'] += replace_across_runs(p, PH_DATE, date)


def unify_captions(root, stats):
    """图/表/算法题注统一：居中、加粗、不缩进；keepNext 按类型设置。"""
    for p in root.iter(q('p')):
        pPr = p.find(q('pPr'))
        if pPr is None:
            continue
        st = pPr.find(q('pStyle'))
        if st is None or st.get(W + 'val') != '5':
            continue
        kind = caption_kind(p)
        if kind is None:
            continue                      # 定义/定理等伪题注，不动
        set_child(pPr, q('jc'), {'val': 'center'}, PPR_ORDER)
        ind = pPr.find(q('ind'))
        if ind is not None:
            for a in ('firstLine', 'firstLineChars', 'left', 'leftChars',
                      'hanging', 'hangingChars'):
                ind.attrib.pop(W + a, None)
            if not ind.attrib:
                pPr.remove(ind)
        for r in visible_runs(p):
            set_child(ensure_rPr(r), q('b'), {}, RPR_ORDER)
        if kind == '图':
            drop_child(pPr, q('keepNext'))
        else:
            set_child(pPr, q('keepNext'), {}, PPR_ORDER)
        stats['captions_fixed'] += 1


def fix_image_keepnext(root, stats):
    """紧邻图题注的图片段落补"与下段同页"。"""
    for container in (root.find(q('body')), ) if root.find(q('body')) is not None else ():
        kids = [el for el in container if el.tag == q('p')]
        for i, p in enumerate(kids):
            if not has_figure(p):
                continue
            nxt = None
            for j in range(i + 1, min(i + 3, len(kids))):
                if ptext(kids[j]).strip():
                    nxt = kids[j]
                    break
            if nxt is None or caption_kind(nxt) != '图':
                continue
            pPr = ensure_pPr(p)
            if pPr.find(q('keepNext')) is None:
                set_child(pPr, q('keepNext'), {}, PPR_ORDER)
                stats['img_keepnext'] += 1


def fill_body_font(root, stats):
    """无样式的长正文段落：补齐 小四 12pt / 宋体。

    跳过表格内的段落：封面的信息表与正文表格都有自己的字号，
    按正文规则覆盖会把它们改坏。
    """
    pm = parent_map(root)
    for p in root.iter(q('p')):
        if pPr_has_table_ancestor(p, pm):
            continue
        pPr = p.find(q('pPr'))
        if pPr is not None and pPr.find(q('pStyle')) is not None:
            continue
        if len(ptext(p).strip()) < MIN_BODY_LEN:
            continue
        if has_figure(p):
            continue
        for r in visible_runs(p):
            rPr = ensure_rPr(r)
            if rPr.find(q('sz')) is None:
                set_child(rPr, q('sz'), {'val': str(SZ_BODY)}, RPR_ORDER)
                set_child(rPr, q('szCs'), {'val': str(SZ_BODY)}, RPR_ORDER)
                stats['font_filled'] += 1
            rf = rPr.find(q('rFonts'))
            if rf is None:
                set_child(rPr, q('rFonts'),
                          {'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman',
                           'eastAsia': '宋体'}, RPR_ORDER)


def strip_comments(z, names, replace, drop):
    """删除批注部件，并修掉 rels 与 Content_Types 中的引用。"""
    parts = [n for n in names if re.match(r'word/comments.*\.xml$', n)
             or n in ('word/people.xml',)]
    for n in parts:
        drop.add(n)
    return parts


def fix_rels(raw, rel_ids_to_drop, target_substrings):
    root = load(raw)
    removed = 0
    for rel in list(root):
        rid = rel.get('Id')
        tgt = rel.get('Target') or ''
        if rid in rel_ids_to_drop or any(s in tgt for s in target_substrings):
            root.remove(rel)
            removed += 1
    return serialize(root, raw), removed


def fix_content_types(raw, part_names):
    root = load(raw)
    removed = 0
    for ov in list(root):
        pn = ov.get('PartName') or ''
        if pn.lstrip('/') in part_names:
            root.remove(ov)
            removed += 1
    return serialize(root, raw), removed


# ---------------------------------------------------------------- 主流程
def prepare(src, dst, course, title=None, date=None, keep_samples=False, verbose=True):
    names = []
    stats = {'red_runs': 0, 'empty_paras': 0, 'rev_removed': 0, 'rev_unwrapped': 0,
             'header_replaced': 0, 'title_replaced': 0, 'date_replaced': 0,
             'captions_fixed': 0, 'img_keepnext': 0,
             'font_filled': 0, 'parts_dropped': 0, 'rels_removed': 0,
             'comment_marks': 0}

    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        replace, drop = {}, set()

        # --- document.xml ---
        raw = z.read('word/document.xml')
        doc = load(raw)
        strip_red_runs(doc, stats)
        accept_revisions(doc, stats)
        if not keep_samples:              # 批注部件要删时，正文锚点必须一起清
            strip_comment_markers(doc, stats)
        fill_placeholders(doc, title, date, stats)
        unify_captions(doc, stats)
        fix_image_keepnext(doc, stats)
        fill_body_font(doc, stats)
        replace['word/document.xml'] = serialize(doc, raw)

        # --- 页眉 / 页脚：只替换占位符 ---
        for n in [x for x in names if re.match(r'word/(header|footer)\d*\.xml$', x)]:
            raw_h = z.read(n)
            h = load(raw_h)
            replace_header_placeholder(h, course, stats)
            replace[n] = serialize(h, raw_h)

        # --- 批注 ---
        if not keep_samples:
            drop.update(strip_comments(z, names, replace, drop))

        # --- document.xml.rels ---
        rels_name = 'word/_rels/document.xml.rels'
        if rels_name in names:
            new_rels, n_removed = fix_rels(
                z.read(rels_name), set(), ['comments', 'people.xml'])
            replace[rels_name] = new_rels
            stats['rels_removed'] = n_removed

        # --- [Content_Types].xml ---
        if drop and '[Content_Types].xml' in names:
            new_ct, n_removed = fix_content_types(z.read('[Content_Types].xml'), drop)
            replace['[Content_Types].xml'] = new_ct
            stats['parts_dropped'] = n_removed

        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                if item.filename in drop:
                    continue
                data = replace.get(item.filename, None)
                if data is None:
                    data = z.read(item.filename)
                out.writestr(item, data)

    if verbose:
        print('已生成干净起点：%s' % dst)
        print('  删除红色说明文字 run      : %d' % stats['red_runs'])
        print('  删除变空的段落            : %d' % stats['empty_paras'])
        print('  展开修订 / 删除修订标记   : %d / %d'
              % (stats['rev_unwrapped'], stats['rev_removed']))
        print('  清除批注锚点              : %d 处' % stats['comment_marks'])
        print('  页眉占位符替换            : %d' % stats['header_replaced'])
        print('  封面报告名称替换          : %d' % stats['title_replaced'])
        print('  封面日期替换              : %d' % stats['date_replaced'])
        print('  题注格式统一              : %d 处' % stats['captions_fixed'])
        print('  图片段落补"与下段同页"    : %d 处' % stats['img_keepnext'])
        print('  正文 run 补小四字号       : %d' % stats['font_filled'])
        print('  删除批注等部件            : %d 个' % stats['parts_dropped'])

        left = []
        if not stats['title_replaced']:
            left.append(PH_TITLE)
        if not stats['date_replaced']:
            left.append(PH_DATE)
        if left:
            print('\n注意：以下占位符仍留在文档里，交付前必须替换：%s' % '、'.join(left))
        print('\n注意：示例插图、示例表格、定义/定理示例、算法示例框仍需人工删除；')
        print('      封面的姓名/学号/学院/专业/年级是固定文字，换人时需要手改。')
        print('      完成后运行 check_report.py 复核。')
    return stats


def take_opt(args, name, default=None):
    """取出 --name value，并把它从参数表里删掉。"""
    if name in args:
        i = args.index(name)
        if i + 1 >= len(args):
            print('%s 后面要跟一个值' % name)
            sys.exit(2)
        val = args[i + 1]
        del args[i:i + 2]
        return val
    return default


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

    args = list(sys.argv[1:])
    keep_samples = '--keepsamples' in args
    args = [a for a in args if a != '--keepsamples']

    course = take_opt(args, '--course', HDR_PLACEHOLDER)
    title = take_opt(args, '--title')
    date = take_opt(args, '--date')

    if len(args) != 2:
        print(__doc__)
        sys.exit(2)
    prepare(args[0], args[1], course, title=title, date=date,
            keep_samples=keep_samples)


if __name__ == '__main__':
    main()

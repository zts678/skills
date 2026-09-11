# -*- coding: utf-8 -*-
"""
排版一致性校验（tcm-operation-spec-docx 技能质量门禁）

用法：
  python verify_layout.py <基准文档.docx> <新文档.docx>
退出码 0 = 封面/正文排版签名全部一致；1 = 存在差异。

校验项：页面尺寸与四向页边距、Normal 样式（宋体sz24）、封面 7 段结构
（字体/字号/加粗/居中/段内换行/段后间距/分页符）、
H1（黑体sz30/段前240段后120）、H2/H3、正文（行距360/lineRule=auto/段后60/首行缩进482，
含"——"引用条目与"[n]"文献条目等短条目）、注释（楷体sz21/灰色595959）。
"""
import sys
from docx import Document
from docx.oxml.ns import qn


def _font_sig(p):
    if not p.runs:
        return None
    rPr = p.runs[0]._element.rPr
    if rPr is None:
        return None
    f = rPr.find(qn('w:rFonts'))
    sz = rPr.find(qn('w:sz'))
    b = rPr.find(qn('w:b'))
    col = rPr.find(qn('w:color'))
    return {
        'font': f.get(qn('w:eastAsia')) if f is not None else None,
        'sz': sz.get(qn('w:val')) if sz is not None else None,
        'bold': b is not None and b.get(qn('w:val')) != '0',
        'color': col.get(qn('w:val')) if col is not None else None,
    }


def _spacing(p):
    pPr = p._p.find(qn('w:pPr'))
    return pPr.find(qn('w:spacing')) if pPr is not None else None


def _breaks(p):
    return [br.get(qn('w:type')) for br in p._p.findall('.//' + qn('w:br'))]


def _centered(p):
    pPr = p._p.find(qn('w:pPr'))
    if pPr is None:
        return False
    jc = pPr.find(qn('w:jc'))
    return jc is not None and jc.get(qn('w:val')) == 'center'


def _body_sig(p):
    pPr = p._p.find(qn('w:pPr'))
    sp = pPr.find(qn('w:spacing')) if pPr is not None else None
    ind = pPr.find(qn('w:ind')) if pPr is not None else None
    return {
        'line': sp.get(qn('w:line')) if sp is not None else None,
        'rule': sp.get(qn('w:lineRule')) if sp is not None else None,
        'after': sp.get(qn('w:after')) if sp is not None else None,
        'firstLine': ind.get(qn('w:firstLine')) if ind is not None else None,
    }


EXPECT_COVER = [
    {'font': '黑体', 'sz': '44', 'bold': True, 'color': None},   # 主标题
    {'font': '楷体', 'sz': '28', 'bold': False, 'color': None},  # 英文副标题（两行）
    {'font': '楷体', 'sz': '26', 'bold': False, 'color': None},  # （操作技术规范）
]


def check(path):
    errs = []
    doc = Document(path)
    sec = doc.sections[0]
    if str(sec.page_width) != '7772400' or str(sec.page_height) != '10058400':
        errs.append('页面尺寸异常: %s x %s' % (sec.page_width, sec.page_height))
    if (str(sec.top_margin), str(sec.bottom_margin)) != ('914400', '914400') or \
            (str(sec.left_margin), str(sec.right_margin)) != ('1143000', '1143000'):
        errs.append('页边距异常: 上%s 下%s 左%s 右%s' % (
            sec.top_margin, sec.bottom_margin, sec.left_margin, sec.right_margin))

    # Normal 样式：宋体 sz24（ascii/hAnsi/eastAsia 三项）
    st_rpr = doc.styles['Normal'].element.find(qn('w:rPr'))
    st_rf = st_rpr.find(qn('w:rFonts')) if st_rpr is not None else None
    st_sz = st_rpr.find(qn('w:sz')) if st_rpr is not None else None
    if st_rf is None:
        errs.append('Normal 样式缺少 rFonts 设置')
    elif st_rf.get(qn('w:eastAsia')) != '宋体' or \
            st_rf.get(qn('w:ascii')) != '宋体' or st_rf.get(qn('w:hAnsi')) != '宋体':
        errs.append('Normal 样式字体异常: %s/%s/%s' % (
            st_rf.get(qn('w:ascii')), st_rf.get(qn('w:hAnsi')), st_rf.get(qn('w:eastAsia'))))
    if st_sz is None or st_sz.get(qn('w:val')) != '24':
        errs.append('Normal 样式字号应为sz24')

    ps = doc.paragraphs

    # 封面 0/1/2
    for i, exp in enumerate(EXPECT_COVER):
        sig = _font_sig(ps[i])
        if sig is None:
            errs.append('封面段[%d]无文字run' % i)
            continue
        for k, v in exp.items():
            if sig[k] != v:
                errs.append('封面段[%d] %s=%r，期望 %r（%s）' % (i, k, sig[k], v, ps[i].text[:20]))
        if not _centered(ps[i]):
            errs.append('封面段[%d]未居中' % i)

    # 主标题段后 120
    sp0 = _spacing(ps[0])
    if sp0 is None or sp0.get(qn('w:after')) != '120':
        errs.append('主标题段后间距应为120')

    # 英文副标题段内换行 + 段后360
    brs = _breaks(ps[1])
    if None not in brs:
        errs.append('英文副标题缺少段内换行(LINE break)')
    sp = ps[1]._p.find(qn('w:pPr')).find(qn('w:spacing'))
    if sp is None or sp.get(qn('w:after')) != '360':
        errs.append('英文副标题段后间距应为360')

    # 空行 x2
    if ps[3].text != '' or ps[4].text != '':
        errs.append('封面第3、4段应为空行')

    # 适用于行
    sig5 = _font_sig(ps[5])
    if not sig5 or sig5['font'] != '宋体' or sig5['sz'] != '24' or not _centered(ps[5]):
        errs.append('"适用于"行格式异常: %r' % sig5)

    # 分页符
    if 'page' not in _breaks(ps[6]):
        errs.append('封面缺少分页符（编制说明应从第2页开始）')

    # 编制说明应为 H1
    sig7 = _font_sig(ps[7])
    if ps[7].text.strip() != '编制说明' or not sig7 or sig7['font'] != '黑体' or sig7['sz'] != '30':
        errs.append('第2页首段应为 H1"编制说明"(黑体sz30)，实际: %s %r' % (ps[7].text, sig7))

    # 全文扫描：H1间距、正文格式与注释格式
    seen_body = seen_note = seen_h2 = seen_h3 = False
    for p in ps:
        t = p.text.strip()
        sig = _font_sig(p)
        if not sig or not t:
            continue
        # 正文：所有非居中的宋体sz24常规段落（含"——"引用、[n]文献等短条目）
        if sig['font'] == '宋体' and sig['sz'] == '24' and not sig['bold'] and not _centered(p):
            bs = _body_sig(p)
            if bs['line'] != '360' or bs['rule'] != 'auto' or bs['after'] != '60' or bs['firstLine'] != '482':
                errs.append('正文格式异常 "%s…": %r' % (t[:12], bs))
            seen_body = True
        if t.startswith('注：'):
            if sig['font'] != '楷体' or sig['sz'] != '21' or sig['color'] != '595959':
                errs.append('注释格式异常: font=%s sz=%s color=%s' % (sig['font'], sig['sz'], sig['color']))
            seen_note = True
        if sig['font'] == '黑体' and sig['sz'] == '30' and sig['bold']:
            sp = _spacing(p)
            if sp is None or sp.get(qn('w:before')) != '240' or sp.get(qn('w:after')) != '120':
                errs.append('H1段间距异常(应before=240/after=120) "%s…"' % t[:12])
        if sig['font'] == '黑体' and sig['sz'] == '26' and sig['bold']:
            seen_h2 = True
        if sig['font'] == '黑体' and sig['sz'] == '24' and sig['bold']:
            seen_h3 = True

    if not seen_body:
        errs.append('未找到正文段落')
    if not seen_note:
        errs.append('未找到"注："注释段')
    if not seen_h2:
        errs.append('未找到 H2 节标题')
    return errs


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    base_errs = check(sys.argv[1])
    new_errs = check(sys.argv[2])
    print('基准文档 %s：%d 项问题' % (sys.argv[1], len(base_errs)))
    for e in base_errs:
        print('  [基准]', e)
    print('新文档   %s：%d 项问题' % (sys.argv[2], len(new_errs)))
    for e in new_errs:
        print('  [新文档]', e)
    if new_errs:
        print('RESULT: FAIL')
        sys.exit(1)
    print('RESULT: PASS（新文档封面与正文排版签名符合基准规范）')
    sys.exit(0)

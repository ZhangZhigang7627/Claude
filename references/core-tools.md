# core-tools.md — ppt_builder.py 完整代码 + 写入规则

## ppt_builder.py（每次任务开始复制到 /home/claude/ppt_builder.py）

```python
"""ppt_builder.py — 基于sp索引的可靠内容写入工具"""
import re

BASE = '/home/claude/unpacked/ppt/slides'  # 按实际任务调整

def esc(t):
    return t.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

def load(slide_num):
    with open(f'{BASE}/slide{slide_num}.xml', 'r', encoding='utf-8') as f:
        return f.read()

def save(slide_num, content):
    """中文输出：替换lang为zh-CN。英文输出：把这行改为替换zh-CN→en-US"""
    content = re.sub(r'lang="en-(?:US|GB)"', 'lang="zh-CN"', content)
    with open(f'{BASE}/slide{slide_num}.xml', 'w', encoding='utf-8') as f:
        f.write(content)

def make_para(text, bold=False, sz=1200, space_before=0, line_spacing=0):
    """正文段落。sz默认1200=12pt。正文必须加line_spacing=150000(1.5倍)。标题传ls=0。"""
    b = ' b="1"' if bold else ''
    spc = f'<a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>'
    lns = f'<a:lnSpc><a:spcPct val="{line_spacing}"/></a:lnSpc>' if line_spacing else ''
    return (f'<a:p><a:pPr>{lns}{spc}<a:buNone/></a:pPr>'
            f'<a:r><a:rPr lang="zh-CN" sz="{sz}"{b} dirty="0">'
            f'<a:latin typeface="Microsoft YaHei"/>'
            f'<a:ea typeface="Microsoft YaHei"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p>')

def make_para_colored(text, color='000000', bold=False, sz=1200, space_before=0,
                      align='ctr', line_spacing=0):
    """有背景填充的形状专用（圆圈/色块/矩形）。必须用此函数否则白字白底看不见。
    fontRef继承lt1(白色)，make_para()写进去的文字会不可见。"""
    b = ' b="1"' if bold else ''
    spc = f'<a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>'
    lns = f'<a:lnSpc><a:spcPct val="{line_spacing}"/></a:lnSpc>' if line_spacing else ''
    return (f'<a:p><a:pPr algn="{align}">{lns}{spc}<a:buNone/></a:pPr>'
            f'<a:r><a:rPr lang="zh-CN" sz="{sz}"{b} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Microsoft YaHei"/>'
            f'<a:ea typeface="Microsoft YaHei"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p>')

def _get_sp_spans(content):
    return [(m.start(), m.end()) for m in re.finditer(r'<p:sp>.*?</p:sp>', content, re.DOTALL)]

def rewrite_sp(content, sp_index, new_paras_xml):
    """核心：整体重写sp的txBody。同时支持a:txBody和p:txBody两种命名空间。"""
    spans = _get_sp_spans(content)
    if sp_index >= len(spans):
        print(f'⚠️ sp[{sp_index}]不存在（共{len(spans)}个）')
        return content
    s, e = spans[sp_index]
    old_sp = content[s:e]
    def replace_txbody(m):
        txbody = m.group(0); tag = m.group(1); close = f'</{tag}:txBody>'
        hm = re.match(rf'(<{tag}:txBody>.*?<a:lstStyle/>)', txbody, re.DOTALL)
        if not hm: hm = re.match(rf'(<{tag}:txBody>.*?</a:lstStyle>)', txbody, re.DOTALL)
        return (hm.group(1) + new_paras_xml + close) if hm else txbody
    new_sp = re.sub(r'<([ap]):txBody>.*?</\1:txBody>', replace_txbody, old_sp, flags=re.DOTALL)
    return content[:s] + new_sp + content[e:]

def map_slide(slide_num):
    """查看所有sp内容和索引，多run的sp必须用rewrite_sp整体重写。"""
    c = load(slide_num); spans = _get_sp_spans(c)
    print(f'=== slide{slide_num} ({len(spans)}sp) ===')
    for i, (s, e) in enumerate(spans):
        runs = re.findall(r'<a:t[^>]*>([^<]*)</a:t>', c[s:e])
        joined = ''.join(runs)
        if joined.strip():
            multi = ' [多run!]' if len([r for r in runs if r.strip()]) > 1 else ''
            print(f'  sp[{i}]{multi}: {repr(joined[:80])}')

def verify(slide_num):
    """写完立即调用，确认全中文无英文残留。两句粘连是显示问题非错误。"""
    c = load(slide_num); spans = _get_sp_spans(c)
    print(f'--- slide{slide_num} verify ---')
    for i, (s, e) in enumerate(spans):
        t = ''.join(re.findall(r'<a:t[^>]*>([^<]*)</a:t>', c[s:e]))
        if t.strip(): print(f'  [{i}] {repr(t[:80])}')

def survey(slide_num):
    """≥20sp的slide写内容前必须调用。按y坐标确认视觉顺序，按cx找内容框。"""
    c = load(slide_num); spans = _get_sp_spans(c)
    print(f'=== slide{slide_num} 坐标普查（{len(spans)}sp）===')
    for i, (s, e) in enumerate(spans):
        sp = c[s:e]
        off = re.search(r'<a:off x="(\d+)" y="(\d+)"', sp)
        ext = re.search(r'<a:ext cx="(\d+)"', sp)
        if not off: continue
        x = int(off.group(1))//914400; y = int(off.group(2))//914400
        cx = int(ext.group(1))//914400 if ext else 0
        t = ''.join(re.findall(r'<a:t[^>]*>([^<]*)</a:t>', sp)).strip()
        side = '右' if x > 6 else '左'
        sp_type = '文字框' if cx > 1 else '图标'
        if t or cx > 1:
            print(f'  [{i:2d}] {side} x={x}" y={y}" cx={cx}" {sp_type}: {repr(t[:40]) if t else "(空)"}')
```

## 字号规范（sz = pt × 100）

| 元素 | sz | pt | 行间距 |
|------|----|----|--------|
| 主标题 | 2400 | 24pt | 无（ls=0）|
| 副标题/模块标题 | 1600 | 16pt | 无 |
| **正文** | **1200** | **12pt** | **ls=150000（必须）** |
| 大数字 | 1800 | 18pt | 无（**禁止超过1800**） |
| 章节标签 | 1000 | 10pt | 无 |
| 脚注 | 900 | 9pt | 无 |

## 标准写入示例

```python
import sys; sys.path.insert(0, '/home/claude')
from ppt_builder import *

# ≥20sp 先普查坐标
survey(75)
map_slide(75)

c = load(75)
c = rewrite_sp(c, 0, make_para('主标题', bold=True, sz=2400))
c = rewrite_sp(c, 1, make_para('副标题', sz=1600))

# 正文：sz=1200 + line_spacing=150000，两者缺一不可
c = rewrite_sp(c, 5,
    make_para('模块标题', bold=True, sz=1600) +
    make_para('第一段正文', sz=1200, line_spacing=150000) +
    make_para('第二段正文', sz=1200, line_spacing=150000, space_before=400))

# 圆圈/色块：必须用make_para_colored
c = rewrite_sp(c, 6, make_para_colored('01', color='000000', bold=True, sz=1400))

c = rewrite_sp(c, 22, make_para(''))  # 不需要的框清空
save(75, c)
verify(75)
```

## 关键规则

- **≥20sp先survey**，按y坐标排序（y小=靠上=先写标题），不按sp索引号
- **cx>1"才是内容框**，cx≤1"通常是图标装饰
- **N个图标=N个make_para()段落**，不能用\n\n代替
- **所有sp都要处理**：不用的框写make_para('')清空
- **verify粘连不是错误**：两句话显示粘连是join问题，不是XML问题
- **p:pic图标元素**：从其他slide复制spTree时，p:pic会引用media文件，clean.py清理后断链，必须删除所有`<p:pic>...</p:pic>`元素

# core-tools.md — ppt_builder.py 完整代码 + 写入规则
<!-- v1.5 · 2026-03-31 · 新增get_ph()按类型定位placeholder；新增del_gf()删除graphicFrame；修复系统性标题顺序陷阱 -->

## ppt_builder.py（每次任务开始复制到 /home/claude/ppt_builder.py）

```python
"""ppt_builder.py v1.5
修复三个系统性问题：
1. get_ph(): 按ph类型定位placeholder，永久解决标题顺序陷阱
2. del_gf(): 删除graphicFrame，永久解决英文表格残留问题
3. save()内置_force_dark_color(): 自动注入333333，解决白字继承问题
"""
import re

BASE = '/home/claude/unpacked/ppt/slides'

def esc(t):
    return t.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

def load(slide_num):
    with open(f'{BASE}/slide{slide_num}.xml', 'r', encoding='utf-8') as f:
        return f.read()

def _force_dark_color(content, default='333333'):
    """强制为所有无显式颜色的rPr注入深色，彻底消灭白字继承问题。"""
    def inject(m):
        rpr = m.group(0)
        if '<a:solidFill>' in rpr: return rpr
        return re.sub(r'(<a:rPr[^>]*>)',
                      r'\1<a:solidFill><a:srgbClr val="' + default + r'"/></a:solidFill>',
                      rpr, count=1)
    return re.sub(r'<a:rPr[^>]*>.*?</a:rPr>', inject, content, flags=re.DOTALL)

def save(slide_num, content):
    """保存slide。自动完成：①lang替换为zh-CN ②注入深色防白字"""
    content = re.sub(r'lang="en-(?:US|GB)"', 'lang="zh-CN"', content)
    content = _force_dark_color(content)
    with open(f'{BASE}/slide{slide_num}.xml', 'w', encoding='utf-8') as f:
        f.write(content)

# ============================================================
# ★ v1.5 新增：get_ph() — 彻底解决标题顺序陷阱
# ============================================================

def get_ph(content):
    """【必须用这个函数定位placeholder，禁止用sp[0]/sp[1]硬编码】
    
    返回 {ph_type: sp_index} 映射，例如：
      {'title': 1, 'body': 0}  ← 这个模板大多数slide都是这个顺序！
    
    根本问题：这个模板68张slide里有64张 sp[0]=body, sp[1]=title
    XML索引顺序与视觉顺序相反，用sp[0]写主标题会写到副标题位置。
    
    正确用法：
      ph = get_ph(c)
      c = rs(c, ph['title'], mph('主标题', bold=True))
      c = rs(c, ph['body'],  mp('副标题', sz=1300))
    """
    sps = re.findall(r'<p:sp>.*?</p:sp>', content, re.DOTALL)
    result = {}
    for i, sp in enumerate(sps):
        m = re.search(r'<p:ph([^/]*)/>', sp)
        if not m: continue
        attrs = m.group(1)
        ph_type_m = re.search(r'type="([^"]+)"', attrs)
        ph_type = ph_type_m.group(1) if ph_type_m else 'body'
        # 记录：用ph类型作key，如果重复取第一个
        if ph_type not in result:
            result[ph_type] = i
        # 也记录body的别名
        if ph_type in ('body', 'subTitle'):
            result.setdefault('body', i)
    return result

# ============================================================
# ★ v1.5 新增：del_gf() — 彻底解决graphicFrame英文表格残留
# ============================================================

def del_gf(content):
    """删除slide里所有graphicFrame（表格/图表）。
    
    根本问题：这个模板有5张slide含graphicFrame（硬编码英文表格/图表）：
      slide08, slide12, slide13, slide22, slide64
    graphicFrame不是<p:sp>，rewrite_sp()无法访问，会原样保留英文内容。
    
    对所有克隆自这些slide的页面，必须在写入内容前调用del_gf()。
    对其他slide调用也安全（无graphicFrame时返回原内容不变）。
    
    用法：
      c = load(slide_num)
      c = del_gf(c)          # ← 先删除graphicFrame
      c = rs(c, ...)         # 再写入内容
      save(slide_num, c)
    """
    return re.sub(r'<p:graphicFrame>.*?</p:graphicFrame>', '', content, flags=re.DOTALL)

# ============================================================
# 段落构造函数（三种）
# ============================================================

def make_ph_para(text, bold=False, space_before=0, line_spacing=0, align='l'):
    """【placeholder专用】不写sz，字号由母版slideLayout决定。
    必须配合 get_ph() 使用，不要手动用sp[0]/sp[1]硬编码索引。
    """
    b = ' b="1"' if bold else ''
    spc = f'<a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>'
    lns = f'<a:lnSpc><a:spcPct val="{line_spacing}"/></a:lnSpc>' if line_spacing else ''
    ag = f' algn="{align}"' if align != 'l' else ''
    return (f'<a:p><a:pPr{ag}>{lns}{spc}<a:buNone/></a:pPr>'
            f'<a:r><a:rPr lang="zh-CN"{b} dirty="0">'
            f'<a:latin typeface="Microsoft YaHei"/>'
            f'<a:ea typeface="Microsoft YaHei"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p>')

# 简写别名（推荐使用）
mph = make_ph_para

def make_para(text, bold=False, sz=1200, space_before=0, line_spacing=0, align='l'):
    """普通内容框段落，写固定sz。save()自动注入333333颜色防白字。"""
    b = ' b="1"' if bold else ''
    spc = f'<a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>'
    lns = f'<a:lnSpc><a:spcPct val="{line_spacing}"/></a:lnSpc>' if line_spacing else ''
    ag = f' algn="{align}"' if align != 'l' else ''
    return (f'<a:p><a:pPr{ag}>{lns}{spc}<a:buNone/></a:pPr>'
            f'<a:r><a:rPr lang="zh-CN" sz="{sz}"{b} dirty="0">'
            f'<a:latin typeface="Microsoft YaHei"/>'
            f'<a:ea typeface="Microsoft YaHei"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p>')

mp = make_para

def make_para_colored(text, color='000000', bold=False, sz=1200, space_before=0,
                      align='ctr', line_spacing=0):
    """显式颜色段落。用于有色背景框的白字，或需要特定品牌色的场景。
    
    ⚠️ 白底描边圆（bg1白色填充+accent3蓝边框）上的编号：
       不要用 color='FFFFFF'（白字白底不可见）
       应用 color='62B5E5'（accent3蓝色，与边框同色，在白底可见）
    """
    b = ' b="1"' if bold else ''
    spc = f'<a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>'
    lns = f'<a:lnSpc><a:spcPct val="{line_spacing}"/></a:lnSpc>' if line_spacing else ''
    return (f'<a:p><a:pPr algn="{align}">{lns}{spc}<a:buNone/></a:pPr>'
            f'<a:r><a:rPr lang="zh-CN" sz="{sz}"{b} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Microsoft YaHei"/>'
            f'<a:ea typeface="Microsoft YaHei"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p>')

mpc = make_para_colored

# ============================================================
# 核心写入函数
# ============================================================

def _get_sp_spans(content):
    return [(m.start(), m.end()) for m in re.finditer(r'<p:sp>.*?</p:sp>', content, re.DOTALL)]

def rewrite_sp(content, sp_index, new_paras_xml):
    """核心：整体重写sp的txBody。支持a:txBody和p:txBody两种命名空间。"""
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

rs = rewrite_sp

# ============================================================
# 辅助工具函数
# ============================================================

def map_slide(slide_num):
    """查看所有sp内容和索引。"""
    c = load(slide_num); spans = _get_sp_spans(c)
    print(f'=== slide{slide_num} ({len(spans)}sp) ===')
    for i, (s, e) in enumerate(spans):
        runs = re.findall(r'<a:t[^>]*>([^<]*)</a:t>', c[s:e])
        joined = ''.join(runs)
        ph = re.search(r'<p:ph([^/]*)/>', c[s:e])
        ph_str = f' [{ph.group(0)}]' if ph else ''
        if joined.strip():
            print(f'  sp[{i}]{ph_str}: {repr(joined[:80])}')

def verify(slide_num):
    """写完立即调用，确认全中文无英文残留。"""
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
        ph = re.search(r'<p:ph([^/]*)/>', sp)
        t = ''.join(re.findall(r'<a:t[^>]*>([^<]*)</a:t>', sp)).strip()
        ph_str = f'[{ph.group(0)}]' if ph else ''
        if t or cx > 1:
            print(f'  [{i:2d}] x={x}" y={y}" cx={cx}" {ph_str}: {repr(t[:40]) if t else "(空)"}')
```

## ★ 三大系统性陷阱速查（v1.5新增）

### 陷阱1：标题顺序陷阱（影响该模板64/68张slide）

**症状**：主标题内容写反了，显示在副标题位置（下方小字）。  
**根因**：该模板几乎所有slide都是 `sp[0]=body`，`sp[1]=title`，XML索引顺序与视觉位置相反。  
**永久修复**：用 `get_ph()` 按类型定位，禁止用 `sp[0]/sp[1]` 硬编码。

```python
# ❌ 旧写法（会写反）
c = rs(c, 0, mph('主标题内容', bold=True))   # sp[0]=body，写到了下方小字框！
c = rs(c, 1, mp('副标题内容', sz=1300))       # sp[1]=title，写到了上方大字框！

# ✅ 新写法（永远正确）
ph = get_ph(c)
c = rs(c, ph['title'], mph('主标题内容', bold=True))   # 永远找到title
c = rs(c, ph['body'],  mp('副标题内容', sz=1300))      # 永远找到body
```

### 陷阱2：graphicFrame英文表格残留（该模板5张slide含graphicFrame）

**症状**：页面显示英文表格内容，或中英文重叠乱掉。  
**受影响的源slide**：slide08, slide12, slide13, slide22, slide64  
**永久修复**：load()之后立即调用 `del_gf()`。

```python
# ✅ 正确写法（对所有slide都安全，无graphicFrame时无副作用）
c = load(slide_num)
c = del_gf(c)          # ← 先删除graphicFrame
ph = get_ph(c)
c = rs(c, ph['title'], mph('主标题', bold=True))
save(slide_num, c)
```

### 陷阱3：白底描边圆上的白字（make_para_colored误用FFFFFF）

**症状**：圆圈编号不可见（白字+白底）。  
**根因**：圆圈是"白底+accent3蓝边框"的描边圆，`mpc(color='FFFFFF')` 写白字在白底上不可见。`_force_dark_color()` 看到已有solidFill就跳过，救不回来。  
**永久修复**：圆圈编号用 `color='62B5E5'`（accent3蓝色），在白底上清晰可见。

```python
# ❌ 错误（白底白字不可见）
c = rs(c, 5, mpc('1', color='FFFFFF', bold=True, sz=1800))

# ✅ 正确（蓝色文字在白底描边圆上可见）
c = rs(c, 5, mpc('1', color='62B5E5', bold=True, sz=1800))
```

**如何判断圆圈类型**（写入前检查spPr）：
```python
# 检查sp[5]是蓝色填充圆还是白底描边圆
import re
with open(f'unpacked/ppt/slides/slideN.xml') as f: c = f.read()
sps = re.findall(r'<p:sp>.*?</p:sp>', c, re.DOTALL)
sp5 = sps[5]
sppr = re.search(r'<p:spPr>(.*?)</p:spPr>', sp5, re.DOTALL).group(1)
fill_rgb = re.findall(r'solidFill.*?srgbClr val="([^"]+)"', sppr, re.DOTALL)
fill_sch = re.findall(r'solidFill.*?schemeClr val="([^"]+)"', sppr, re.DOTALL)
ln_sch   = re.findall(r'<a:ln.*?schemeClr val="([^"]+)"', sppr, re.DOTALL)
print(f"填充: rgb={fill_rgb} scheme={fill_sch}")
print(f"描边: scheme={ln_sch}")
# 如果填充=bg1(白) + 描边=accent3(蓝) → 白底描边圆 → 用color='62B5E5'
# 如果填充=accent3/accent1(有色) → 有色填充圆 → 用color='FFFFFF'
```

## 字号规范（sz = pt × 100）

| 元素 | 函数 | sz | 备注 |
|------|------|-----|------|
| **主标题 placeholder** | **`mph()` + `get_ph()['title']`** | **不写** | **母版决定，禁止用sp[0/1]硬编码** |
| **副标题 placeholder** | **`mph()` + `get_ph()['body']`** | **不写** | **母版决定，禁止用sp[0/1]硬编码** |
| 普通标题/模块标题 | `mp()` | 1600 | 非placeholder的内容框 |
| **正文** | **`mp()`** | **1200** | **ls=150000（必须）** |
| 大数字 | `mp()` | 1800 | 禁止超过1800 |
| 白底描边圆编号 | `mpc(color='62B5E5')` | 1600-1800 | 蓝色文字在白底可见 |
| 有色填充圆编号 | `mpc(color='FFFFFF')` | 1600-1800 | 白色文字在深色背景可见 |

## 标准写入示例（v1.5完整版）

```python
import sys; sys.path.insert(0, '/home/claude')
from ppt_builder import *

# ====== 示例1：普通内容页（双栏Goals版式）======
c = load(289)
c = del_gf(c)          # 先删graphicFrame（安全操作）
ph = get_ph(c)         # 获取placeholder索引映射

# ✅ 用get_ph()写placeholder，永远正确
c = rs(c, ph['title'], mph('体系一：现状分析', bold=True))
c = rs(c, ph['body'],  mp('初步建成矩阵式组织、制度框架与方法论', sz=1300))

# 普通内容框用索引（非placeholder，顺序是稳定的）
c = rs(c, 3, mpc('已建成', bold=True, sz=1300))
c = rs(c, 2,
    mp('▶ 战略与组织', bold=True, sz=1200, line_spacing=150000) +
    mp('正文内容', sz=1100, line_spacing=150000, space_before=100))

save(289, c)
verify(289)

# ====== 示例2：体系标题页（slide44类型）======
c = load(288)
c = del_gf(c)
ph = get_ph(c)

c = rs(c, ph['title'], mph('数据管理体系', bold=True))
c = rs(c, ph['body'],  mp('从"柔性协调"向"铁腕穿透"的权力升级', sz=1600))
# sp[2]/sp[3]等非placeholder框照常用索引
c = rs(c, 2, mp('正文说明内容', sz=1200, line_spacing=150000))

save(288, c)

# ====== 示例3：三列版式（slide17带白底描边圆）======
c = load(294)
c = del_gf(c)
ph = get_ph(c)

c = rs(c, ph['title'], mph('体系二：现状分析', bold=True))
c = rs(c, ph['body'],  mp('完成底数静态盘点与架构初建', sz=1300))

# ✅ 白底描边圆：用accent3蓝色(62B5E5)，在白底可见
c = rs(c, 5, mpc('1', color='62B5E5', bold=True, sz=1800))
c = rs(c, 6, mpc('2', color='62B5E5', bold=True, sz=1800))
c = rs(c, 7, mpc('3', color='62B5E5', bold=True, sz=1800))

save(294, c)
```

## 关键规则（v1.5更新）

- **【v1.5新增】始终用 `get_ph()` 定位placeholder**，禁止硬编码 `sp[0]`/`sp[1]`
- **【v1.5新增】始终在 load() 后立即调用 `del_gf()`**，对所有slide无副作用
- **【v1.5新增】圆圈编号先检查类型**：白底描边圆用 `62B5E5`，有色填充圆用 `FFFFFF`
- **≥20sp先survey**，按y坐标排序（y小=靠上=先写标题），不按sp索引号
- **cx>1"才是内容框**，cx≤1"通常是图标装饰
- **N个图标=N个make_para()段落**，不能用\n\n代替
- **所有sp都要处理**：不用的框写make_para('')清空

---

## ★ v1.5 新增：表格(graphicFrame)写入工具

### 表格结构说明

`<p:graphicFrame>` 里包含 `<a:tbl>` 表格，结构如下：
```
<p:graphicFrame>
  <p:xfrm>...</p:xfrm>          ← 位置和大小
  <a:graphic>
    <a:graphicData>
      <a:tbl>
        <a:tblGrid>              ← 列定义（各列宽度）
          <a:gridCol w="..."/>
        </a:tblGrid>
        <a:tr h="...">           ← 行（row_idx=0,1,2...）
          <a:tc>                 ← 单元格（col_idx=0,1,2...）
            <a:txBody>...</a:txBody>   ← 内容在这里
            <a:tcPr>...</a:tcPr>       ← 单元格样式（边框/填充）
          </a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

### 两种处理策略

| 场景 | 策略 | 函数 |
|------|------|------|
| 表格布局合适，只需换成中文内容 | **保留并写入**（推荐） | `rewrite_tc()` |
| 表格布局不符合需求，不需要表格 | **删除重建** | `del_gf()` + `rs()` |

### rewrite_tc() — 表格单元格内容写入

```python
def make_tc_para(text, bold=False, sz=1200, color='333333', align='l'):
    """表格单元格段落。注意：用a:tcPr样式，不是p:spPr。"""
    def esc(t): return t.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
    b = ' b="1"' if bold else ''
    ag = f' algn="{align}"' if align != 'l' else ''
    return (f'<a:p><a:pPr{ag}><a:buNone/></a:pPr>'
            f'<a:r><a:rPr lang="zh-CN" sz="{sz}"{b} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Microsoft YaHei"/>'
            f'<a:ea typeface="Microsoft YaHei"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r></a:p>')

mtp = make_tc_para  # 简写别名

def rewrite_tc(content, row_idx, col_idx, new_para_xml):
    """重写graphicFrame表格中指定单元格（行row_idx，列col_idx）的内容。
    
    row_idx: 行索引，从0开始（0=标题行，1=第一数据行...）
    col_idx: 列索引，从0开始
    new_para_xml: 用make_tc_para()生成的段落XML
    
    用法示例（slide13的双列原则表格）：
      c = load(310)
      # 不要del_gf()，保留表格，直接写内容
      c = rewrite_tc(c, 0, 1, mtp('核心原则', bold=True, sz=1400))     # 标题行左列
      c = rewrite_tc(c, 0, 3, mtp('对客户的意义', bold=True, sz=1400)) # 标题行右列
      c = rewrite_tc(c, 1, 1, mtp('原则一内容', sz=1200))              # 数据行左列
      c = rewrite_tc(c, 1, 3, mtp('意义一内容', sz=1200))              # 数据行右列
      save(310, c)
    """
    gf_m = re.search(r'(<p:graphicFrame>.*?</p:graphicFrame>)', content, re.DOTALL)
    if not gf_m:
        print("⚠️ 未找到graphicFrame，请先检查slide是否含表格")
        return content
    
    gf_orig = gf_m.group(0)
    rows = list(re.finditer(r'<a:tr[^>]*>.*?</a:tr>', gf_orig, re.DOTALL))
    if row_idx >= len(rows):
        print(f"⚠️ row_idx={row_idx}超出范围（共{len(rows)}行）")
        return content
    
    row_m = rows[row_idx]
    row_str = row_m.group(0)
    cells = list(re.finditer(r'<a:tc>.*?</a:tc>', row_str, re.DOTALL))
    if col_idx >= len(cells):
        print(f"⚠️ col_idx={col_idx}超出范围（共{len(cells)}列）")
        return content
    
    cell_m = cells[col_idx]
    cell_orig = cell_m.group(0)
    
    def replace_txbody(m):
        txbody = m.group(0)
        hm = re.match(r'(<a:txBody>.*?<a:lstStyle/>)', txbody, re.DOTALL)
        if not hm: hm = re.match(r'(<a:txBody>.*?</a:lstStyle>)', txbody, re.DOTALL)
        if hm: return hm.group(1) + new_para_xml + '</a:txBody>'
        return txbody
    
    cell_new = re.sub(r'<a:txBody>.*?</a:txBody>', replace_txbody, cell_orig, flags=re.DOTALL)
    row_new = row_str[:cell_m.start()] + cell_new + row_str[cell_m.end():]
    gf_new = gf_orig[:row_m.start()] + row_new + gf_orig[row_m.end():]
    return content[:gf_m.start()] + gf_new + content[gf_m.end():]

rtc = rewrite_tc  # 简写别名

def scan_table(slide_num):
    """查看graphicFrame表格结构（行列数和内容），用于写入前规划。"""
    c = load(slide_num)
    gf = re.search(r'<p:graphicFrame>.*?</p:graphicFrame>', c, re.DOTALL)
    if not gf:
        print(f"slide{slide_num}无graphicFrame")
        return
    gf_str = gf.group(0)
    rows = re.findall(r'<a:tr[^>]*>.*?</a:tr>', gf_str, re.DOTALL)
    cols = re.findall(r'<a:gridCol w="(\d+)"', gf_str)
    print(f"=== slide{slide_num}表格: {len(rows)}行×{len(cols)}列 ===")
    print(f"列宽(英寸): {[round(int(w)/914400,2) for w in cols]}")
    for ri, row in enumerate(rows):
        cells = re.findall(r'<a:tc>.*?</a:tc>', row, re.DOTALL)
        print(f"\n  行{ri}:")
        for ci, cell in enumerate(cells):
            texts = re.findall(r'<a:t[^>]*>([^<]*)</a:t>', cell)
            text = ''.join(texts).strip()
            print(f"    列{ci}: {repr(text[:50]) if text else '(空)'}")
```

### 该模板5个graphicFrame的具体结构

| 源slide | 表格结构 | 典型用途 | 推荐策略 |
|---------|---------|---------|---------|
| slide08 | 2列×多行（受访者列表） | 访谈名单 | 通常del_gf |
| slide12 | 多列流程 | 流程步骤 | 可rewrite_tc |
| **slide13** | **4列×7行（左文字/右文字）** | **原则+说明双列表** | **推荐rewrite_tc** |
| slide22 | 复杂矩阵 | 功能对比 | 可rewrite_tc |
| slide64 | 2列（before/after） | 对比表格 | 推荐rewrite_tc |

### slide13双列原则表格完整写入示例

```python
# slide13是最常用的带图标双列原则表格
# 结构：4列（图标 | 左文字 | 分隔 | 右文字），第1/3列是内容列
c = load(slide_num)
ph = get_ph(c)
c = rs(c, ph['title'], mph('体系一：业务深水区痛点', bold=True))
c = rs(c, ph['body'],  mp('悬浮的管控与失效的指挥棒', sz=1300))

# 写表格（保留图标和样式，只替换文字）
c = rtc(c, 0, 1, mtp('核心痛点', bold=True, sz=1400))       # 标题行：左列标题
c = rtc(c, 0, 3, mtp('业务代价', bold=True, sz=1400))       # 标题行：右列标题
c = rtc(c, 1, 1, mtp('"管理套管理"的陷阱：面对18家下属公司，集团数字化部门沦为没有实权的"数据汇总员"', sz=1100))
c = rtc(c, 1, 3, mtp('管控指令在层层下达中被"软抵制"，形成悬浮于业务之上的表层管理', sz=1100))
c = rtc(c, 2, 1, mtp('MTP考核不与真实业务痛点挂钩', sz=1100))
c = rtc(c, 2, 3, mtp('基层数据管家"干好干坏一个样"，治理工作沦为应付检查的"走过场"', sz=1100))
c = rtc(c, 3, 1, mtp('资源保障机制缺失', sz=1100))
c = rtc(c, 3, 3, mtp('预算安排与"十四五"规划匹配度不够，数据工作高度依赖依附性项目预算', sz=1100))
# 清空多余行（如果数据行少于模板行数）
c = rtc(c, 4, 1, mtp('')); c = rtc(c, 4, 3, mtp(''))
c = rtc(c, 5, 1, mtp('')); c = rtc(c, 5, 3, mtp(''))
c = rtc(c, 6, 1, mtp('')); c = rtc(c, 6, 3, mtp(''))
save(slide_num, c)
```

### 何时用 rewrite_tc vs del_gf

```python
# ✅ 用 rewrite_tc：表格布局适合内容，保留美观的样式和图标
# 例如：slide13的双列表格，保留左侧图标圆圈，替换文字
c = load(slide_num)
ph = get_ph(c)
c = rs(c, ph['title'], mph('标题', bold=True))
c = rtc(c, 0, 1, mtp('左列标题', bold=True))
# ... 

# ✅ 用 del_gf：表格内容与新需求完全不匹配，不需要表格版式
# 例如：克隆slide13但要做完全不同布局的内容
c = load(slide_num)
c = del_gf(c)        # 删除表格
ph = get_ph(c)
c = rs(c, ph['title'], mph('标题', bold=True))
c = rs(c, 3, mp('用普通sp写内容', sz=1200))
save(slide_num, c)
```

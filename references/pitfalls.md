# pitfalls.md — 污染清理 + 常见问题速查 + 血泪教训
<!-- v1.5 · 2026-03-31 · 新增血泪教训五（graphicFrame系统性盲区）；更新血泪教训三/四为已修复状态 -->

## think-cell 污染清理

```python
import zipfile, re

with zipfile.ZipFile('/home/claude/template.pptx', 'r') as z:
    bad = ['THINKCELL','think-cell','TCLayout','oleObj','AlternateContent','custDataLst']
    for name in z.namelist():
        if not (name.endswith('.xml') or name.endswith('.rels')): continue
        c = z.read(name).decode('utf-8', errors='replace')
        hits = [k for k in bad if k.lower() in c.lower()]
        if hits: print(f"  {name}: {hits}")
```

（完整清理脚本见v1.3版本，此处略）

## 污染位置速查表

| 位置 | 污染内容 | 正确处理 |
|------|---------|---------|
| ppt/tags/tag*.xml | THINKCELL... | **替换为空壳**（不删文件！） |
| ppt/slides/slide*.xml | custDataLst+oleObj+AlternateContent | 删除这些XML元素 |
| ppt/slides/_rels/*.rels | oleObject关系引用 | 删除对应Relationship行 |
| ppt/embeddings/oleObject*.bin | OLE二进制数据 | 删除文件+更新Content_Types |

## 常见问题速查

| 现象 | 根因 | 解决 |
|------|------|------|
| 主标题内容在下方、副标题在上方 | **见血泪教训四**（v1.5已修复） | 用`get_ph()`定位，禁止sp[0/1]硬编码 |
| 页面出现英文表格 / 中英文重叠 | **见血泪教训五**（v1.5已修复） | load()后立即调用`del_gf()` |
| 文字写进去但看不见（白字） | **见血泪教训三**（v1.4已修复） | v1.4 save()自动注入333333 |
| 圆圈编号不可见 | **见血泪教训三补充**（v1.5已修复） | 白底描边圆用`color='62B5E5'` |
| PowerPoint报"修复并删除" | broken references | 检查rels里是否有引用但文件不存在 |
| 英文残留（最常见） | 多run碎片，字符串替换静默失败 | map_slide找出，rewrite_sp整体重写 |
| 标题内容上下颠倒 | 按sp索引号写，但sp[0].y>sp[1].y | 用`get_ph()`，不用索引号 |
| 右侧/底部空白 | 未做坐标普查，漏写cx=5"的大文字框 | survey()找cx最大的sp |
| 大数字版式失调 | sz>1800在双栏版式里比例失调 | 大数字最大sz=1800 |
| sldIdLst幻灯片数量虚报 | re.findall扫描整个pres.xml误匹配slideMaster | 只扫描sldIdLst内部的rId |
| 会话结束文件丢失 | /home/claude/是临时目录 | 质检通过后立即cp到/mnt/user-data/outputs/ |

---

## 血泪教训一：schemeClr 颜色漂移（变紫/变蓝）

**现象**：视觉检查发现背景色变成紫色或蓝色。  
**解决**：从模板实际`theme.xml`动态读取颜色值，硬编码替换schemeClr，禁止手填固定值。

---

## 血泪教训二：xfrm 注入导致标题错位

**现象**：视觉检查发现标题飞出页面。  
**解决**：placeholder（title/body）的xfrm由slideLayout继承，**永远不要手动注入xfrm**。

---

## 血泪教训三：白底白字——QA检测不到的隐形bug（v1.4已修复）

**现象**：文字不可见（白字白底），但QA脚本报告无问题。  
**根因**：`theme.xml → lt1 = FFFFFF`，通过fontRef隐式传递，rPr里无显式FFFFFF，QA永远检测不到。  
**v1.4修复**：`save()` 内置 `_force_dark_color()`，自动为无solidFill的rPr注入333333。

**血泪教训三补充：make_para_colored误用FFFFFF（v1.5已修复）**  
`_force_dark_color()` 只能救"没有颜色"的情况；如果主动写了 `color='FFFFFF'`，它看到已有solidFill就跳过，FFFFFF被保留，仍然白字白底。  
- 白底描边圆（bg1白色填充+accent3蓝边框）：用 `color='62B5E5'`
- 有色填充圆（accent1/accent2等有色填充）：用 `color='FFFFFF'`

---

## 血泪教训四：placeholder字号被覆盖（v1.4已修复）

**现象**：主标题字号小于副标题，字号层级颠倒。  
**根因**：`make_para(sz=2400)` 覆盖了母版的字号定义。  
**v1.4修复**：提供 `make_ph_para()`（不写sz），让字号由slideLayout继承。

---

## 血泪教训五：graphicFrame系统性盲区（v1.5已修复）

**现象A**：页面出现英文表格内容（如"KEY PRINCIPLES / IMPLICATIONS FOR THE CLIENT"）。  
**现象B**：中文内容与英文表格重叠，页面布局乱掉。  

**根因**（两层）：  
1. **该模板有5张slide含graphicFrame**（硬编码英文表格/图表）：`slide08, slide12, slide13, slide22, slide64`
2. **graphicFrame是 `<p:graphicFrame>` 元素，不是 `<p:sp>`**，所以：
   - `rewrite_sp()` 完全无法访问它
   - `_force_dark_color()` 不处理它的颜色
   - 语言残留检测（检查`<a:t>`）只看sp，看不到graphicFrame里的文字
   - 结果：graphicFrame被原样保留，英文内容一直在那里

**侦查方法**：在侦查阶段批量检测所有候选slide是否含graphicFrame：
```python
import re
with open(f'unpacked/ppt/slides/slideN.xml') as f: c = f.read()
gfs = re.findall(r'<p:graphicFrame>', c)
if gfs:
    gf_texts = re.findall(r'<a:t[^>]*>([^<]+)</a:t>',
                          ''.join(re.findall(r'<p:graphicFrame>.*?</p:graphicFrame>', c, re.DOTALL)))
    print(f"⚠️ slideN含{len(gfs)}个graphicFrame: {[t for t in gf_texts if t.strip()][:5]}")
```

**v1.5修复**：在 `ppt_builder.py` 中提供 `del_gf()` 函数，在每次 `load()` 后立即调用：
```python
c = load(slide_num)
c = del_gf(c)   # ← 删除所有graphicFrame，对无graphicFrame的slide无副作用
ph = get_ph(c)
c = rs(c, ph['title'], mph('主标题', bold=True))
save(slide_num, c)
```

**该模板5张含graphicFrame的slide**（选版式时优先避开，或clone后必须del_gf）：
| slide | graphicFrame内容 |
|-------|----------------|
| slide08 | Company Interviews / We Quals / Other（访谈表格） |
| slide12 | Identify / Capture / Embed（流程表格） |
| slide13 | KEY PRINCIPLES / IMPLICATIONS FOR THE CLIENT（双列原则表格） |
| slide22 | Self-Service and Seamless eCommerce...（功能矩阵） |
| slide64 | BEFORE / AFTER（对比表格） |

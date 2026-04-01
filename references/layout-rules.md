# layout-rules.md — 版式选择 + 坐标规则 + 黄金标准

## sp普查脚本（第四步）

```python
import re, os

base = 'unpacked'
slide_files = sorted(
    [f for f in os.listdir(f'{base}/ppt/slides')
     if f.startswith('slide') and f.endswith('.xml') and '_rels' not in f],
    key=lambda x: int(re.search(r'\d+', x).group()))

print(f'{"slide":>10} {"sp":>5}  {"档":6}  首行文字')
for fname in slide_files:
    snum = int(re.search(r'\d+', fname).group())
    with open(f'{base}/ppt/slides/{fname}') as f: c = f.read()
    n = len(re.findall(r'<p:sp>', c))
    tier = '○极简' if n<=6 else ('·标准' if n<=22 else '★丰富')
    spans = [(m.start(),m.end()) for m in re.finditer(r'<p:sp>.*?</p:sp>',c,re.DOTALL)]
    title = next((''.join(re.findall(r'<a:t[^>]*>([^<]*)</a:t>',c[s:e])).strip()[:30]
                  for s,e in spans[:2]
                  if ''.join(re.findall(r'<a:t[^>]*>([^<]*)</a:t>',c[s:e])).strip()), '')
    print(f'  slide{snum:<4} {n:4d}sp {tier}  {"█"*(n//3)}  {title}')
```

## 版式多样性硬性要求

- **sp差值 ≥ 20**：最大值-最小值必须≥20
- **三档都要有**：极简(≤6sp) + 标准(7-22sp) + 丰富(≥25sp)
- 章节片段场景下三档都要有，但可以没有极简页（封面/谢谢）

版式类型速查：

| 视觉类型 | sp数量区间 | 典型用途 |
|---------|-----------|---------|
| 3列卡片 | 20-25sp | 三项并列成果 |
| Goals双栏 | 28-33sp | 优势vs差距、现状vs目标 |
| 5图标横排 | 28-35sp | 五项并列痛点/举措 |
| 带勾列表 | 7-10sp | 4-5条带图标列表 |
| 流程横排(before/after) | 25-32sp | 改革路径、路线图 |
| 双栏+数据块 | 30-40sp | 痛点+量化数据 |
| 3列编号卡片 | 10-14sp | 三大举措+编号 |
| 矩阵/全景图 | 45sp+ | 复杂对比/全景总结 |

## 选源slide + 更新sldIdLst

### add_slide.py 行为说明（必读，避免双重注册）

`add_slide.py` **每次只接受一个 source 参数**。调用后它会自动完成：
- ✅ 复制 slide XML 文件到 unpacked/ppt/slides/
- ✅ 注册到 `ppt/_rels/presentation.xml.rels`
- ✅ 注册到 `[Content_Types].xml`
- ❌ **不会**自动写入 `presentation.xml` 的 sldIdLst

**禁止**在调用 `add_slide.py` 之后再手动往 rels 注册同一个 slide——那会导致双重注册，输出文件页数翻倍。

### 正确流程（必须按顺序执行）

```python
import subprocess, re

# Step 1: 对每个源slide各调用一次 add_slide.py
# 它完成：文件复制 + rels注册 + Content_Types注册
sources = [42, 11, 45, 41, 17, 37]  # 根据规划调整
ORIGINAL_SLIDE_COUNT = 68           # ⚠️ 原模板实际slide数，必须按实际填写

for src in sources:
    r = subprocess.run(
        ['python3', '/mnt/skills/public/pptx/scripts/add_slide.py', 'unpacked/', f'slide{src}.xml'],
        capture_output=True, text=True)
    print(r.stdout.strip())  # 确认每个slide的新编号和分配到的rId

# Step 2: 从rels里找出所有新增slide（编号 > 原模板最大编号）
with open('unpacked/ppt/_rels/presentation.xml.rels') as f: rels = f.read()

new_slides = sorted(
    [(int(re.search(r'slide(\d+)', m.group(0)).group(1)), m.group(1))
     for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="slides/(slide(\d+))\.xml"', rels)
     if int(re.search(r'slide(\d+)', m.group(0)).group(1)) > ORIGINAL_SLIDE_COUNT],
    key=lambda x: x[0])

print(f'检测到 {len(new_slides)} 个新slide: {new_slides}')

# Step 3: 完整替换 sldIdLst — 只写入新页，原模板页全部踢出
# ⚠️ 这是输出文件页数正确的关键：替换而非追加
entries = [f'<p:sldId id="{700+i}" r:id="{rid}"/>'
           for i, (snum, rid) in enumerate(new_slides)]
new_lst = '<p:sldIdLst>\n        ' + '\n        '.join(entries) + '\n    </p:sldIdLst>'

with open('unpacked/ppt/presentation.xml') as f: pres = f.read()
pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', new_lst, pres, flags=re.DOTALL)
with open('unpacked/ppt/presentation.xml', 'w') as f: f.write(pres)
print(f'✅ sldIdLst已替换，输出 {len(new_slides)} 张（原模板页已全部移除）')
```

## 坐标判断规则

| 情况 | 判断 | 处理 |
|------|------|------|
| 找内容框 | cx最大的sp才是内容区，cx≤1"通常是图标 | 只向cx>1"的sp写文字 |
| 视觉顺序 | y小=靠上=标题，y大=靠下=内容 | 按y坐标排序，不按sp索引号 |
| 左右分栏 | x>6"为右侧，x≤6"为左侧 | 内容分别写入对应侧sp |
| 图标数=段落数 | N个图标必须N个make_para() | 不能用\n\n代替段落分隔 |

## 黄金标准：好PPT vs 差PPT

**叙事结构**（最根本）：
- ❌差：体系一→体系二→体系三，每个套相同模板
- ✅好：现状痛点(为什么) → 核心方案(怎么做) → 具体战役(先做什么)，每页有结论

**页数控制**：
- 汇报场景15-20页最佳，超过25页必须审查冗余
- 宁可一页讲3件事，不要3页各讲1件事

**每页内容**：
- 副标题是对标题的一句话论证（不是重复标题）
- 每页底部/右侧必须有"核心结论"框，一句话提炼本页观点
- 相邻两页要有逻辑承接，不是独立孤岛

**模板残留检查**：选源slide后，p:pic图标元素会引用media文件，被clean.py清理后成断链。
凡从spTree复制的p:pic元素，必须检查并删除：
```python
import re
with open(f'unpacked/ppt/slides/slide{n}.xml') as f: c = f.read()
pics = re.findall(r'<p:pic>.*?</p:pic>', c, re.DOTALL)
print(f'slide{n}: {len(pics)}个p:pic元素')  # 有的话必须删除
c = re.sub(r'<p:pic>.*?</p:pic>', '', c, flags=re.DOTALL)
with open(f'unpacked/ppt/slides/slide{n}.xml', 'w') as f: f.write(c)
```

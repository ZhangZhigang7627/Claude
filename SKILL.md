---
name: ppt-template-editor
description: >
  【优先级最高，覆盖所有PPT相关任务】以下任何情况都必须使用本Skill，严禁使用ppt-generator：
  1. 用户上传了.pptx格式的模板文件；
  2. 用户说"继续上次的PPT任务"、"继续任务"、"读取session文件继续"——即使本次对话没有上传.pptx，
     只要session文件或历史记录中提到了.pptx模板，就必须用本Skill；
  3. 用户说"用这个模板做PPT"、"按模板样式生成"、"在模板上填内容"；
  4. 任务上下文中存在.pptx模板路径（如/mnt/user-data/uploads/*.pptx）。
  【技术原因】ppt-template-editor通过复制模板slide保留完整母版引用链（slide→slideLayout→slideMaster→theme），
  生成的PPT自动继承模板的背景/配色/字体/图形元素。
  ppt-generator从零创建，没有母版引用链，生成的PPT永远丢失母版样式。有模板就绝不能用ppt-generator。
  核心原则：复用.pptx模板的版式和样式，只替换文字内容，质检通过后必须输出到/mnt/user-data/outputs/。
---

# ppt-template-editor

输入：内容文档（Word/MD）+ PPTX模板 → 输出：高质量多页PPT  
支持全量生成（5-35页）和增量修改（单页/几页）。

---

## ⚠️ 第0步（每次必做，执行前先确认）

收到任务后，按以下顺序执行，**不要直接开始生成**：

**0.1 声明使用的 Skill**

先回复用户：
> 我将使用 **`ppt-template-editor`** 来执行这个任务。  
> 原因：`[一句话说明原因]`  
> 请确认是否继续？

等待用户确认后进入 0.2。

**0.2 检查必要文件是否存在**

```python
import os, glob

uploads = '/mnt/user-data/uploads'

# 检查 .pptx 模板文件
pptx_files = glob.glob(f'{uploads}/*.pptx')
# 检查 Word 内容文件
docx_files = glob.glob(f'{uploads}/*.docx')

missing = []
if not pptx_files: missing.append('❌ 未找到 .pptx 模板文件')
if not docx_files: missing.append('❌ 未找到 .docx 内容文件')

if missing:
    print("缺少以下必要文件，请上传后继续：")
    for m in missing: print(m)
    print("\n需要：")
    print("  1. PPTX 模板文件（.pptx）")
    print("  2. PPT 内容文件（.docx）")
    # 停止执行，等待用户上传
else:
    print(f"✅ 模板文件: {pptx_files}")
    print(f"✅ 内容文件: {docx_files}")
```

如果文件缺失，**停止执行，明确告知用户需要上传哪个文件**，不要继续后续步骤。

---

## ⚠️ 七条铁律（违反即失败）

1. **只用 `rewrite_sp()` 写内容**，禁止字符串匹配替换（多run文本会静默失败）
2. **有背景填充的形状用 `make_para_colored(color='000000')`**，永远不用 `FFFFFF`——白字白底 QA 检测不出来
3. **质检通过后立即 `cp` 到 `/mnt/user-data/outputs/`**，`/home/claude/` 会话结束即清空
4. **新会话开始先读 session 文件**：`/mnt/user-data/outputs/session_任务名.md`
5. **sldIdLst 必须完整替换，只保留新页**：输出文件只含本次生成的N页，原模板所有页必须从 sldIdLst 移除。见第5步。
6. **打包前必须硬编码 schemeClr**：否则 PowerPoint 颜色漂移变紫色。见 `references/pitfalls.md` 血泪教训一。
7. **打包前必须注入 title/body placeholder xfrm 到 `<p:spPr>` 内部**：否则标题位置在 PowerPoint 里错乱。见 `references/pitfalls.md` 血泪教训二。

---

## 标准流程（9步，Step 0 在所有步骤之前）

| 步骤 | 操作 | 参考文档 |
|------|------|---------|
| **0. 确认Skill + 检查文件** | 声明Skill → 等确认 → 检查文件存在 → 缺失则停止 | 见上方第0步 |
| **1. 确认需求** | PPT范围/语言/页数/排版规范/叙事主线 | 见下方第一步 |
| **2. 创建session文件** | 写入任务状态到outputs/session_xxx.md | 见下方session格式 |
| **3. 污染清理** | 检测并清理think-cell污染 | `references/pitfalls.md` |
| **4. 解包模板** | unpack.py → 生成缩略图 → sp普查 | `references/layout-rules.md` |
| **5. 规划版式** | 选源slide，sp差值≥20，三档齐全 | `references/layout-rules.md` |
| **6. 写入内容** | ppt_builder工具库，逐页survey→RS→verify | `references/core-tools.md` |
| **7. 清理打包** | ① schemeClr硬编码 ② xfrm注入 ③ clean.py ④ pack.py | `references/pitfalls.md` 血泪教训一/二 |
| **8. 质检** | 8.1结构 8.2污染 8.3残留 8.4颜色 8.4b白字白底 8.5多样 8.6视觉 | `references/qa-scripts.md` |
| **9. 输出** | cp到outputs + present_files + 更新session | 见下方 |

---

## 第一步：确认需求（必须逐项确认）

```
1. PPT范围：完整独立 / 章节片段（无封面目录） / 单页增量
2. 模板语言 → 输出语言：英文→中文 / 英文→英文 / 中文→中文 / 中文→英文
3. 页数目标：精简(5-15) / 标准(15-25) / 完整(25-40)
4. 排版规范（从模板提取或用户指定）：
   默认：字体=Microsoft YaHei, 主标题sz=2400, 副标题sz=1600,
         正文sz=1200+line_spacing=150000(1.5倍), 大数字sz=1800(上限)
5. 叙事主线（2-3句话）：为什么→怎么做→先做什么
```

语言决策：
- 中文输出：save() 替换 lang="zh-CN"，字体 Microsoft YaHei
- 英文输出：保留 lang="en-US"，字体保留模板原字体
- 残留检测：中文输出检测英文残留；英文输出检测中文残留；英文→英文跳过

---

## 第二步：创建/读取 session 文件

**新任务开始**：创建 `/mnt/user-data/outputs/session_任务名.md`  
**新会话继续**：第一件事读 session 文件，从断点继续

> ⚠️ **重要**：新会话说"继续上次的PPT任务"时，即使用户没有重新上传 `.pptx` 文件，
> 也必须使用本Skill（ppt-template-editor），**绝不能**切换到 ppt-generator。
> 原因：ppt-template-editor 通过复制模板 slide 保留完整母版引用链；
> ppt-generator 从零创建，生成的 PPT 没有母版，样式全部丢失。
> 模板文件路径在 session 文件里，从 `/mnt/user-data/uploads/` 读取即可。

```
# session_任务名.md 格式

## 任务信息
- 模板：/mnt/user-data/uploads/xxx.pptx（已清理版：template_clean.pptx）
- 内容：/mnt/user-data/uploads/xxx.docx
- 输出：/mnt/user-data/outputs/xxx.pptx
- 语言：英文模板→中文输出
- 排版：字体YaHei, 正文sz=1200+ls=150000, 主标题sz=2400

## 版式规划
| 页 | 源slide | sp | 内容 | 状态 |
|---|---|---|---|---|
| 1 | slide42 | 22sp | 四项建设成果 | 已完成 |
| 2 | slide11 | 31sp | 政策响应度 | 进行中 |

## 当前状态
- 已完成：1确认 2session 3污染清理 4解包 5规划 6写入(第1页)
- 下一步：写入第2页 slide70

## 输出文件
- （质检通过后填写路径）
```

每完成一个关键步骤立即更新 session 文件的"当前状态"。

---

## 核心命令速查

```bash
# 解包
python3 /mnt/skills/public/pptx/scripts/office/unpack.py template_clean.pptx unpacked/
# 缩略图
python3 /mnt/skills/public/pptx/scripts/thumbnail.py template.pptx thumb --cols 5
# 复制源slide（⚠️ 每次只传一个slide；自动完成：文件复制+rels注册+Content_Types注册；
#              但【不会】自动写入 presentation.xml 的 sldIdLst）
python3 /mnt/skills/public/pptx/scripts/add_slide.py unpacked/ slideN.xml
# 清理
python3 /mnt/skills/public/pptx/scripts/clean.py unpacked/
# 打包
python3 /mnt/skills/public/pptx/scripts/office/pack.py unpacked/ output.pptx --original template_clean.pptx
# 高清QA图
python3 /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf /home/claude/slide
pdftoppm -jpeg -r 150 -f N -l M output.pdf /home/claude/slide  # 只生成第N-M页
```

### ⚠️ 第5步完成后必须执行：sldIdLst 完整替换

```python
import re, subprocess

sources = [42, 11, 45, 41, 17, 37]   # ← 按版式规划填写
ORIGINAL_SLIDE_COUNT = 68             # ← 原模板实际slide总数，必须按实际填写

for src in sources:
    r = subprocess.run(
        ['python3', '/mnt/skills/public/pptx/scripts/add_slide.py', 'unpacked/', f'slide{src}.xml'],
        capture_output=True, text=True)
    print(r.stdout.strip())

with open('unpacked/ppt/_rels/presentation.xml.rels') as f: rels = f.read()
new_slides = sorted(
    [(int(re.search(r'slide(\d+)', m.group(0)).group(1)), m.group(1))
     for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="slides/(slide(\d+))\.xml"', rels)
     if int(re.search(r'slide(\d+)', m.group(0)).group(1)) > ORIGINAL_SLIDE_COUNT],
    key=lambda x: x[0])

entries = [f'<p:sldId id="{700+i}" r:id="{rid}"/>' for i,(snum,rid) in enumerate(new_slides)]
new_lst = '<p:sldIdLst>\n        ' + '\n        '.join(entries) + '\n    </p:sldIdLst>'
with open('unpacked/ppt/presentation.xml') as f: pres = f.read()
pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', new_lst, pres, flags=re.DOTALL)
with open('unpacked/ppt/presentation.xml', 'w') as f: f.write(pres)
print(f'✅ sldIdLst已替换，输出 {len(new_slides)} 张（原模板页已全部移除）')
```

### ⚠️ 第7步必须按顺序执行：schemeClr硬编码 → xfrm注入 → clean → pack

读取 `references/pitfalls.md` 血泪教训一和二获取完整代码。

```
① hardcode_scheme_colors() — 对所有新slide + slideMaster + slideLayout
② inject_title_xfrm()      — 对所有新slide的sp[0](body13) 和 sp[1](title)
③ python3 clean.py unpacked/
④ python3 pack.py unpacked/ output.pptx --original template_clean.pptx
```

```python
# 第九步输出（必须执行）
import shutil
shutil.copy('/home/claude/output.pptx', '/mnt/user-data/outputs/最终文件名.pptx')
# 然后调用 present_files 工具，再更新 session 文件
```

---

## 参考文档索引（按需读取，不要全量加载）

| 文档 | 内容 | 何时读 |
|------|------|--------|
| `references/core-tools.md` | ppt_builder.py完整代码 + 写入规则 | 开始写内容前读一次 |
| `references/layout-rules.md` | sp普查 + 版式选择 + 坐标判断 + 黄金标准 | 规划版式时 |
| `references/qa-scripts.md` | 全量质检脚本（8.1-8.6六项） | 质检时 |
| `references/pitfalls.md` | 污染清理 + 常见问题速查 + 血泪教训 | 遇到报错或异常时 |
| `references/incremental.md` | 增量修改四种操作 + 快速验证 | 单页修改时 |

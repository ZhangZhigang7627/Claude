# ppt-template-editor

> **优先级：最高** — 当涉及 `.pptx` 模板时，所有 PPT 相关任务必须使用本 Skill。

通过复用已有 `.pptx` 模板的版式和样式（仅替换文字内容）来生成或修改 PPT 的技能。

---

## 功能 | What It Does

**输入：** 内容文档（Word `.docx` / Markdown）+ PPTX 模板  
**输出：** 高质量多页 PPT（全量生成 5–35 页，或单页增量修改）

| 方案 | 原理 | 效果 |
|---|---|---|
| **ppt-template-editor** ✅ | 复制模板 slide，保留完整母版链：`slide → slideLayout → slideMaster → theme` | 背景、配色、字体、图形元素自动继承 |
| **ppt-generator** ❌ | 从零创建，无母版引用链 | 样式永远丢失——有模板时绝不能用 |

---

## 何时使用 | When to Use

满足以下**任意条件**时，立即使用本 Skill：

1. ✅ 用户上传了 `.pptx` 模板文件
2. ✅ 用户说"继续上次的 PPT 任务"、"继续任务"、"读取 session 继续"——即使本次对话没有重新上传 `.pptx`，session 文件中也会引用模板路径
3. ✅ 用户说"用这个模板做 PPT"、"按模板样式生成"、"在模板上填内容"
4. ✅ 任务上下文中存在 `.pptx` 模板路径（如 `/mnt/user-data/uploads/*.pptx`）

---

## 9 步流程 | The 9-Step Workflow

| 步骤 | 操作 | 参考文档 |
|------|--------|-----------|
| **0. 确认** | 声明 Skill → 检查文件 → 缺失则停止 | 上方第 0 步 |
| **1. 确认需求** | 范围 / 语言 / 页数 / 排版规范 / 叙事主线 | 见下方第一步 |
| **2. 创建 session 文件** | 写入任务状态到 `outputs/session_xxx.md` | 见下方 session 格式 |
| **3. 污染清理** | 检测并清理 think-cell 污染 | `references/pitfalls.md` |
| **4. 解包模板** | `unpack.py` → 缩略图 → sp 普查 | `references/layout-rules.md` |
| **5. 规划版式** | 选源 slide，sp 差值 ≥ 20，三档齐全 | `references/layout-rules.md` |
| **6. 写入内容** | ppt_builder 工具库：逐页 survey → RS → verify | `references/core-tools.md` |
| **7. 清理打包** | ① 硬编码 schemeClr → ② 注入 xfrm → ③ clean.py → ④ pack.py | `references/pitfalls.md` 血泪教训 1&2 |
| **8. 质检** | 8.1 结构 · 8.2 污染 · 8.3 残留 · 8.4 颜色 · 8.4b 白字白底 · 8.5 多样 · 8.6 视觉 | `references/qa-scripts.md` |
| **9. 输出** | cp 到 `outputs/` + 更新 session | 见下方 |

---

## ⚠️ 七条铁律 | Seven Iron Rules（违反即失败）

1. **只用 `rewrite_sp()` 写内容** — 字符串匹配替换对多 run 文本会静默失败
2. **有背景填充的形状用 `make_para_colored(color='000000')`** — 永远不要用 `FFFFFF`（白字白底质检无法检测）
3. **质检通过后立即 cp 到 `/mnt/user-data/outputs/`** — `/home/claude/` 会话结束即清空
4. **新会话开始先读 session 文件** — `outputs/session_任务名.md`
5. **`sldIdLst` 必须完整替换，只保留新页** — 输出文件只含本次生成的 N 页，原模板所有页必须从 sldIdLst 移除。见第 5 步。
6. **打包前必须硬编码 schemeClr** — 否则 PowerPoint 颜色漂移变紫色。见 `references/pitfalls.md` 血泪教训一。
7. **打包前必须注入 title/body placeholder xfrm 到 `<p:spPr>` 内部** — 否则标题位置在 PowerPoint 里错乱。见 `references/pitfalls.md` 血泪教训二。

---

## 第一步 — 确认需求 | Step 1 — Confirm Requirements

```
1. PPT 范围 / Scope:    完整独立 / 章节片段（无封面目录） / 单页增量
2. 语言 / Language:      模板 → 输出: 英文→中文 / 英文→英文 / 中文→中文 / 中文→英文
3. 页数目标 / Pages:     精简(5-15) / 标准(15-25) / 完整(25-40)
4. 排版规范 / Layout:    （从模板提取或用户指定）
   默认：字体=Microsoft YaHei, 主标题sz=2400, 副标题sz=1600,
         正文sz=1200 + line_spacing=150000 (1.5倍), 大数字sz=1800 (上限)
5. 叙事主线 / Narrative: （2-3句话：为什么→怎么做→先做什么）
```

语言规则 / Language Rules:
- **中文输出**：在 `save()` 中替换 `lang="zh-CN"`，字体 = Microsoft YaHei
- **英文输出**：保留 `lang="en-US"`，保留模板原字体
- **残留检测**：中文输出检测英文残留；英文输出检测中文残留；英文→英文跳过此检测

---

## 第二步 — Session 文件格式 | Step 2 — Session File Format

**新任务开始：** 创建 `/mnt/user-data/outputs/session_任务名.md`  
**继续任务：** 第一件事读 session 文件，从断点继续

```
## 任务信息 / Task Info
- 模板 Template: /mnt/user-data/uploads/xxx.pptx（已清理版 Cleaned: template_clean.pptx）
- 内容 Content: /mnt/user-data/uploads/xxx.docx
- 输出 Output: /mnt/user-data/outputs/xxx.pptx
- 语言 Language: 英文模板 → 中文输出 (EN template → ZH output)
- 排版 Layout: 字体=YaHei, 正文sz=1200 + ls=150000, 主标题sz=2400

## 版式规划 / Layout Plan
| 页 Page | 源slide Source | sp | 内容 Content | 状态 Status |
|---------|---------------|----|-------------|------------|
| 1       | slide42       | 22 | 四项建设成果 Four achievements | 已完成 Done |
| 2       | slide11       | 31 | 政策响应度 Policy responsiveness | 进行中 In progress |

## 当前状态 / Current Status
- 已完成 Completed: 1确认 · 2session · 3污染清理 · 4解包 · 5规划 · 6写入(第1页)
- 下一步 Next: 写入第2页 slide70

## 输出文件 / Output
- （质检通过后填写 / fill in after QA passes）
```

每完成一个关键步骤立即更新 session 文件。

---

## 核心命令 | Core Commands

```bash
# 解包 / Unpack
python3 /mnt/skills/public/pptx/scripts/office/unpack.py template_clean.pptx unpacked/
# 缩略图 / Thumbnails
python3 /mnt/skills/public/pptx/scripts/thumbnail.py template.pptx thumb --cols 5
# 复制源slide（一次只传一个；自动完成：文件复制+rels注册+Content_Types注册；但【不会】写入 sldIdLst）
python3 /mnt/skills/public/pptx/scripts/add_slide.py unpacked/ slideN.xml
# 清理 / Clean
python3 /mnt/skills/public/pptx/scripts/clean.py unpacked/
# 打包 / Pack
python3 /mnt/skills/public/pptx/scripts/office/pack.py unpacked/ output.pptx --original template_clean.pptx
# 高清质检图 / QA Screenshots (HD)
python3 /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf /home/claude/slide
pdftoppm -jpeg -r 150 -f N -l M output.pdf /home/claude/slide  # 只生成第N-M页 / pages N to M only
```

### 第 5 步完成后 — sldIdLst 完整替换 | After Step 5 — sldIdLst Full Replacement

```python
import re, subprocess

sources = [42, 11, 45, 41, 17, 37]   # ← 按版式规划填写 / fill in per layout plan
ORIGINAL_SLIDE_COUNT = 68             # ← 原模板实际总数 / must match actual count

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

### 第 7 步必须按顺序执行 | Step 7 Sequence (Must Follow This Order)

读取 `references/pitfalls.md` 血泪教训一和二获取完整代码 / Read Lessons 1 & 2 for full code.

```
① hardcode_scheme_colors() — 对所有新 slide + slideMaster + slideLayout
② inject_title_xfrm()      — 对所有新 slide 的 sp[0] (body13) 和 sp[1] (title)
③ python3 clean.py unpacked/
④ python3 pack.py unpacked/ output.pptx --original template_clean.pptx
```

### 第 9 步输出（必须执行）| Step 9 — Output (Mandatory)

```python
import shutil
shutil.copy('/home/claude/output.pptx', '/mnt/user-data/outputs/最终文件名.pptx')
# 然后调用 present_files，再更新 session 文件
```

---

## 参考文档 | Reference Documents

| 文档 Document | 内容 Content | 何时读 When to Read |
|----------|---------|-------------|
| `references/core-tools.md` | ppt_builder.py 完整代码 + 写入规则 | 开始写内容前读一次 |
| `references/layout-rules.md` | sp 普查 + 版式选择 + 坐标判断 + 黄金标准 | 规划版式时 |
| `references/qa-scripts.md` | 全量质检脚本（8.1–8.6 六项） | 质检时 |
| `references/pitfalls.md` | 污染清理 + 常见问题速查 + 血泪教训 | 遇到报错或异常时 |
| `references/incremental.md` | 增量修改四种操作 + 快速验证 | 单页修改时 |

---

## 快速决策流程图 | Quick Decision Flowchart

```
用户请求 PPT 任务 / User requests PPT task?
├─ 是 YES → 有 .pptx 模板涉及吗？/ Is a .pptx template involved?
│         ├─ 是 YES → 使用 ppt-template-editor ✅
│         └─ 否 NO  → 用户说了"继续"且 session 文件引用了 .pptx？
│                   └─ 是 YES → 使用 ppt-template-editor ✅
├─ 否 NO  → 从零创建全新 PPT（无模板）？
          └─ 是 YES → 使用 ppt-generator（独立 Skill）
```

---

## 技术背景 | Tech Background

**为什么 ppt-template-editor 能保留样式：**

PowerPoint 模板有层级结构：

```
slide（页）
  └─ slideLayout（版式/母版页）
       └─ slideMaster（母版slide）
            └─ theme（配色、字体、效果）
```

从模板复制 slide 时，整个引用链被保留。PowerPoint 在渲染时通过引用链解析样式——背景、配色、字体、图形元素自动继承。

从零创建 slide 则没有母版链。PowerPoint 没有引用可供解析，样式丢失。

**结论：有模板时必用 template editor；仅对无模板任务使用 ppt-generator。**
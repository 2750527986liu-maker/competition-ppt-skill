---
name: competition-ppt
description: 中国国际大学生创新大赛（国赛）路演PPT制作。当用户提到"竞赛PPT"、"国赛"、"路演PPT"、"创新创业大赛"、"比赛幻灯片"时触发。基于两份国赛金奖级模板的设计方法论，使用python-pptx + PIL生成原创高质量PPT。
---

# 竞赛PPT制作 Skill

## 触发条件
用户需要制作参加中国国际大学生创新大赛（或类似创新创业比赛）的路演PPT。

## 第一步：需求收集（必须先做）

在开始制作前，用 AskUserQuestion 工具收集以下信息：

1. **项目文件**：用户是否有策划书、路演稿、流程大纲等文件？如果有，先全部读取理解内容。
2. **参考模板**：用户是否提供了参考PPT模板？如果有，必须用python-pptx分析其结构（见"模板分析"章节），**不要直接复制模板素材**，而是学习其设计思路。
3. **页数要求**：一般20-43页，根据内容量决定。
4. **配色偏好**：默认深蓝+金色，用户可指定。

## 第二步：模板分析（如有参考模板）

如果用户提供了参考PPT，**必须**执行以下分析流程：

### 2.1 用python-pptx提取结构信息
```python
from pptx import Presentation
prs = Presentation("模板路径.pptx")
slides_list = list(prs.slides)
print(f"尺寸: {prs.slide_width/914400:.1f} x {prs.slide_height/914400:.1f} 英寸")
print(f"总页数: {len(slides_list)}")

for i, slide in enumerate(slides_list):
    pics = sum(1 for s in slide.shapes if s.shape_type == 13)
    groups = sum(1 for s in slide.shapes if s.shape_type == 6)
    # 提取文本和图片位置
    for shape in slide.shapes:
        if shape.has_text_frame:
            print(f"  文本: {shape.text[:100]} @({shape.left/914400:.1f},{shape.top/914400:.1f})")
        if shape.shape_type == 13:
            print(f"  图片: @({shape.left/914400:.1f},{shape.top/914400:.1f}) {shape.width/914400:.1f}x{shape.height/914400:.1f}")
```

### 2.2 提取背景图片（理解配色和视觉风格）
```python
from pptx.oxml.ns import qn
import os
os.makedirs("template-bg", exist_ok=True)
for i, slide in enumerate(slides_list):
    bg_elem = slide.background._element
    nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
             'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    blip = bg_elem.find('.//a:blipFill/a:blip', nsmap)
    if blip is not None:
        embed = blip.get(qn('r:embed'))
        if embed:
            rel = slide.part.rels[embed]
            img_part = rel.target_part
            ext = img_part.content_type.split('/')[-1].replace('jpeg','jpg')
            with open(f"template-bg/slide{i+1:02d}.{ext}", 'wb') as f:
                f.write(img_part.blob)
```

### 2.3 用Read工具查看背景图片
将提取的背景图片复制到用户可见的工作文件夹，用Read工具查看，理解：
- 封面的配色方案（渐变色、主色调）
- 内容页的背景风格（浅色/深色、有无纹理）
- 特殊页面（时间线、致谢页）的设计

### 2.4 总结设计语言（不复制素材）
从模板中学习以下设计原则，然后**自己制作**设计素材：
- 配色方案（主色、辅色、强调色）
- 排版节奏（标题位置、内容区域划分）
- 元素密度（每页多少图片、多少文字块）
- 装饰手法（边框、分割线、图标风格）

## 第三步：内容大纲

### 3.1 国赛标准结构（参考两份模板的共同模式）

根据对两份国赛级PPT（41页冰雪经济 + 43页低空经济）的分析，标准结构如下：

| 编号 | 页面类型 | 内容 | 建议页数 |
|------|---------|------|---------|
| 1 | 封面 | 项目名称+副标题+团队+赛道+负责人 | 1页 |
| 2 | 政策背景 | 国家/行业政策支撑 | 1-2页 |
| 3 | 行业现状 | 市场数据、图表 | 1-2页 |
| 4 | 痛点分析 | 行业核心痛点（通常3个） | 1-2页 |
| 5 | 项目概述 | 技术原理+核心突破+产品展示 | 1-2页 |
| 6 | 研发历程 | 时间线展示 | 1页 |
| 7-N | 核心技术 | 每个核心技术独立1-2页 | 3-6页 |
| N+1 | 核心子系统 | 系统架构展示 | 1-2页 |
| N+2 | 核心母系统 | 整体系统集成 | 1页 |
| N+3 | 衍生技术 | 技术延伸应用 | 1-2页 |
| N+4 | 衍生场景 | 应用场景拓展 | 1页 |
| N+5 | 技术参数 | 关键指标对比表 | 1页 |
| N+6 | 技术壁垒/专利 | 知识产权保护 | 1页 |
| N+7 | 项目荣誉 | 获奖/认证 | 1页 |
| N+8 | 应用案例 | 实际落地案例（每个1页） | 2-4页 |
| N+9 | 应用场景 | 典型使用场景 | 1-2页 |
| N+10 | 竞品分析 | 对比表格/雷达图 | 1页 |
| N+11 | 服务客群/商业模式 | 用户画像+盈利模式 | 2-4页 |
| N+12 | 行业/技术认可 | 合作方背书 | 1-2页 |
| N+13 | 媒体报道 | 权威媒体引用 | 1页 |
| N+14 | 领导关怀 | 领导视察/指导 | 1页（可选） |
| N+15 | 个人成长 | 负责人成长历程 | 1-2页 |
| N+16 | 团队介绍 | 负责人+成员+顾问 | 2-3页 |
| N+17 | 学科引领/带动教育 | 产教融合成果 | 1-2页 |
| N+18 | 社会价值 | 产业+人才+行业价值 | 1页 |
| N+19 | 财务分析 | 收入预测+盈利模型 | 1页 |
| N+20 | 未来规划 | 发展路线图 | 1页 |
| N+21 | 致谢 | 感谢聆听 | 1页 |

**关键原则**：
- 竞赛PPT要**信息密度高、论证充分**，不要怕页数多
- 每个核心技术点**独立成页**
- 团队/教育/社会价值需要**多页展开**
- 案例、荣誉、媒体报道都是**加分页**

### 3.2 页面标题体系（必须遵守）

每页内容页必须包含：
1. **Section标签**：左上角深色圆角矩形，显示页面类别（如"核心技术"、"痛点分析"）
2. **副标题**：以"——"开头，描述本页核心论点
3. **页码**：右下角 "X/总数"

```
[Section标签]  ——副标题描述本页核心论点
─────────────────────────────────────────
              正文内容区域
```

## 第四步：设计素材制作

### 4.1 用PIL生成原创背景（不要复制模板素材）

```python
from PIL import Image, ImageDraw, ImageFilter
import math, random

W, H = 2000, 1125  # 13.33x7.5英寸 @ 150DPI

def gradient(draw, w, h, c1, c2, direction='vertical'):
    for i in range(h if direction == 'vertical' else w):
        ratio = i / (h if direction == 'vertical' else w)
        r = int(c1[0] + (c2[0] - c1[0]) * ratio)
        g = int(c1[1] + (c2[1] - c1[1]) * ratio)
        b = int(c1[2] + (c2[2] - c1[2]) * ratio)
        if direction == 'vertical':
            draw.line([(0, i), (w, i)], fill=(r, g, b))
        else:
            draw.line([(i, 0), (i, h)], fill=(r, g, b))
```

**必须生成的素材**（每种都需要）：
1. **cover-bg.png** (2000x1125)：封面背景。深色渐变+抽象几何/光线/网络线条
2. **content-bg.png** (2000x1125)：内容页背景。浅色+微妙网格/纹理+角落点缀
3. **timeline-bg.png** (2000x1125)：时间线/研发历程背景。彩墨飞溅/渐变色块
4. **dark-section-bg.png** (2000x1125)：深色强调页背景。深蓝+对角线纹理
5. **thankyou-bg.png** (2000x1125)：致谢页背景。深色+中央金色光晕
6. **gold-line.png** (900x6)：金色渐变分割线
7. **section-badge.png** (400x80)：深色section标签底图（可选，也可以用代码画）
8. **moon-icon.png** 或其他项目相关图标 (200x200)

**封面背景制作要点**：
- 深色渐变（如深蓝→更深蓝）
- 添加径向光晕（金色/蓝色）
- 抽象几何线条（网络节点、数据流、网格线）
- 不能太花哨，要保持专业感

**内容页背景制作要点**：
- 浅色为主（白/浅灰/浅蓝）
- 极其微妙的网格纹理
- 角落淡淡的蓝色水彩点缀
- 保持文字可读性

### 4.2 配色方案

默认配色（可按项目主题调整）：
```python
NAVY = RGBColor(0x0A, 0x16, 0x28)        # 深海军蓝，用于标题和section标签
DARK_BLUE = RGBColor(0x0F, 0x1D, 0x35)   # 深蓝
GOLD = RGBColor(0xC9, 0xA8, 0x4C)        # 金色，用于强调和装饰
LIGHT_GOLD = RGBColor(0xE8, 0xD5, 0x8C)  # 浅金色
WHITE = RGBColor(0xFF, 0xFF, 0xFF)        # 白色
DARK_TEXT = RGBColor(0x1E, 0x29, 0x3B)    # 深色正文
SUB_TEXT = RGBColor(0x64, 0x74, 0x8B)     # 灰色副文本
ACCENT_BLUE = RGBColor(0x38, 0x6F, 0xC4) # 蓝色强调
```

### 4.3 字体

- **标题**：等线（DengXian），Windows自带，现代清秀
- **正文**：等线（DengXian）或微软雅黑
- **数据/数字**：等线 Bold，突出显示
- **英文**：等线自动适配

## 第五步：PPT生成代码模板

### 5.1 基础框架

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

W = Inches(13.33)
H = Inches(7.5)
IMG = "项目图片目录"
BG = "设计素材目录"
OUT = "输出路径.pptx"

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]  # blank layout

# 颜色定义（根据项目主题调整）
NAVY = RGBColor(0x0A, 0x16, 0x28)
GOLD = RGBColor(0xC9, 0xA8, 0x4C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT = RGBColor(0x1E, 0x29, 0x3B)
SUB_TEXT = RGBColor(0x64, 0x74, 0x8B)
ACCENT = RGBColor(0x38, 0x6F, 0xC4)
```

### 5.2 核心Helper函数

```python
def add_bg(slide, bg_path):
    """添加全出血背景图"""
    if os.path.exists(bg_path):
        slide.background.fill.background()
        pic = slide.shapes.add_picture(bg_path, Emu(0), Emu(0), W, H)
        sp = pic._element
        sp.getparent().remove(sp)
        slide.shapes._spTree.insert(2, sp)

def text_box(slide, left, top, width, height, text, font_size=18,
             color=DARK_TEXT, bold=False, align=PP_ALIGN.LEFT,
             font_name="等线", line_spacing=1.5):
    """添加文本框"""
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.name = font_name
    if line_spacing:
        p.line_spacing = Pt(font_size * line_spacing)
    return txBox

def multi_text(slide, left, top, width, height, lines, font_size=16,
               color=DARK_TEXT, bold=False, align=PP_ALIGN.LEFT,
               font_name="等线", line_spacing=1.5):
    """添加多段文本"""
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line_text in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(6)
        run = p.add_run()
        run.text = line_text
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
        run.font.bold = bold
        run.font.name = font_name
        p.line_spacing = Pt(font_size * line_spacing)
    return txBox

def gold_line(slide, left, top, width):
    """金色分割线"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(left), Inches(top), Inches(width), Pt(3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = GOLD
    shape.line.fill.background()

def section_badge(slide, text, left=0.3, top=0.2):
    """左上角Section标签"""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(left), Inches(top), Inches(2.6), Inches(0.6))
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY
    shape.line.fill.background()
    tf = shape.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(22)
    run.font.color.rgb = WHITE
    run.font.bold = True
    run.font.name = "等线"

def sub_header(slide, text, left=3.2, top=0.2):
    """副标题（——前缀）"""
    text_box(slide, left, top, 9.5, 0.6, f"——{text}", font_size=16, color=SUB_TEXT)

def card_shape(slide, left, top, width, height, fill=WHITE):
    """白色卡片"""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = RGBColor(0xE2, 0xE8, 0xF0)
    shape.line.width = Pt(1)
    return shape

def add_img(slide, path, left, top, width=None, height=None):
    """安全添加图片"""
    if not os.path.exists(path):
        return None
    if width and height:
        return slide.shapes.add_picture(path, Inches(left), Inches(top),
                                         Inches(width), Inches(height))
    elif width:
        return slide.shapes.add_picture(path, Inches(left), Inches(top), width=Inches(width))
    else:
        return slide.shapes.add_picture(path, Inches(left), Inches(top))

def page_num(slide, num, total):
    """页码"""
    text_box(slide, 12.0, 6.9, 1.0, 0.4, f"{num}/{total}",
             font_size=10, color=SUB_TEXT, align=PP_ALIGN.RIGHT)
```

### 5.3 标准内容页模板

```python
def content_slide(prs, bg_path, badge_text, subtitle, num, total):
    """创建标准内容页骨架"""
    s = prs.slides.add_slide(BLANK)
    add_bg(s, bg_path)
    section_badge(s, badge_text)
    sub_header(s, subtitle)
    page_num(s, num, total)
    return s
```

### 5.4 常用页面布局模式

**三栏卡片布局**（适合数据对比、文献展示）：
```python
for i, (title, desc, img_path) in enumerate(items):
    x = 0.5 + i * 4.2
    card_shape(s, x, 1.3, 3.8, 5.5)
    add_img(s, img_path, x+0.2, 1.5, width=3.4, height=2.8)
    text_box(s, x+0.2, 4.4, 3.4, 0.4, title, font_size=18, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    gold_line(s, x+1.0, 4.9, 1.8)
    text_box(s, x+0.2, 5.1, 3.4, 1.5, desc, font_size=13, color=DARK_TEXT, align=PP_ALIGN.CENTER)
```

**左右分栏布局**（适合产品介绍）：
```python
# 左侧：文字描述
card_shape(s, 0.5, 1.3, 6.0, 5.5)
text_box(s, 0.8, 1.5, 5.4, 0.5, "标题", font_size=28, color=NAVY, bold=True)
gold_line(s, 0.8, 2.1, 3)
multi_text(s, 0.8, 2.3, 5.4, 4.0, bullet_points, font_size=15)
# 右侧：图片
card_shape(s, 6.8, 1.3, 5.8, 2.8)
add_img(s, img_path, 7.0, 1.4, width=5.4, height=2.5)
```

**时间线布局**（适合研发历程）：
```python
for i, (date, title, desc) in enumerate(timeline):
    x = 0.5 + i * 3.15
    # 日期圆形
    circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+0.8), Inches(1.5), Inches(1.2), Inches(1.2))
    circle.fill.solid()
    circle.fill.fore_color.rgb = GOLD
    circle.line.fill.background()
    tf = circle.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = date
    run.font.size = Pt(14)
    run.font.color.rgb = NAVY
    run.font.bold = True
    # 下方内容卡片
    card_shape(s, x, 2.9, 2.9, 3.5)
    text_box(s, x+0.2, 3.0, 2.5, 0.4, title, font_size=16, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    text_box(s, x+0.2, 3.5, 2.5, 2.5, desc, font_size=12, color=DARK_TEXT, align=PP_ALIGN.CENTER)
```

**对比表格布局**（适合竞品分析）：
```python
headers = ["维度", "竞品A", "竞品B", "本项目"]
for i, h in enumerate(headers):
    x = 0.5 + i * 3.15
    shape = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1.3), Inches(2.95), Inches(0.6))
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY if i == 0 else (ACCENT if i == len(headers)-1 else RGBColor(0x64,0x74,0x8B))
    shape.line.fill.background()
    tf = shape.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = h
    run.font.size = Pt(14)
    run.font.color.rgb = WHITE
    run.font.bold = True
```

## 第六步：常见错误与规避

### ❌ 错误1：直接复制模板素材
- **问题**：模板素材有版权，且可能风格不匹配
- **正确做法**：用PIL自己生成背景和装饰元素，学习配色和布局思路

### ❌ 错误2：图片没有固定尺寸
- **问题**：图片大小不一，排版混乱
- **正确做法**：所有图片都指定 width 和 height，保持一致

### ❌ 错误3：每页文字太多
- **问题**：大段文字堆砌，评委看不清重点
- **正确做法**：每页3-5个要点，用关键词+短句，大号字体

### ❌ 错误4：页数太少
- **问题**：20页以下无法充分论证
- **正确做法**：国赛标准30-45页，每个技术点独立成页

### ❌ 错误5：字体不统一
- **问题**：混合使用多种字体，看起来不专业
- **正确做法**：全篇统一使用"等线"（标题加粗，正文正常）

### ❌ 错误6：图片路径包含中文空格
- **问题**：python-pptx可能无法正确加载
- **正确做法**：将图片复制到英文路径的目录中，或用os.path.join确保路径正确

### ❌ 错误7：Write工具截断大文件
- **问题**：用Write工具写大Python文件时可能被截断
- **正确做法**：用bash执行Python来写文件，或分段写入

### ❌ 错误8：中文引号导致语法错误
- **问题**：Python字符串中的""（中文引号）导致SyntaxError
- **正确做法**：所有中文字符串用反引号模板字面量或确保引号匹配

## 第七步：输出与验证

### 7.1 保存并检查文件大小
```bash
ls -la output.pptx
# 正常范围：5-50MB（取决于嵌入图片数量）
# 如果 < 1MB，说明图片没有正确嵌入
```

### 7.2 呈现给用户
使用 `present_files` 工具将PPT呈现给用户，并提供 `computer://` 链接。

### 7.3 迭代优化
用户可能会要求调整：
- 字体大小/类型
- 图片位置/大小
- 内容增减
- 配色修改
- 页数增减

每次修改后重新生成并呈现。

## 第八步：保存记忆

完成PPT后，将以下信息保存到memory：
- 用户的项目名称和类型
- 用户对设计的偏好（配色、字体、风格）
- 制作过程中遇到的问题和解决方案
- 用户确认满意的设计参数（如页数、布局模式）

## 附录：文案写作规律（基于两份国赛PPT文本分析）

### A. 页面文案通用规范

**副标题句式**（每页"——"开头的副标题）：
- "破解...与...两大难题"
- "从...到...的完全闭环"
- "多项数据重大突破，达到国内领先水平！"
- "以项目为熔炉，实现五大核心能力跨越式提升"
- 包含数字增加说服力，如"五大能力"、"三大核心"

**数据引用规范**：
- 每个数据页必须标注来源：`数据来源：中商产业研究院`
- 政府/研究院/协会数据优先
- 标注具体年份增加时效性
- 使用数学符号量化：`XX<数值`、`XX≥数值`、`误差<X%`

### B. 各页面文案模板

#### 封面（第1页）
```
中国国际大学生创新大赛（20XX）
China international College Students' Innovation Competition 20XX
[项目名称]
[一句话定位语，如"以XXX为核心，打造XXX体系，制胜全球XXX竞争"]
[赛道] [负责人姓名]
```

#### 政策背景（第2页）
```
section标签: 政策背景
副标题: ——[概括政策对行业的推动作用]

[机构名称] · [年月]《[政策全称]》
[一句话概括该政策如何推动本行业]
```
**规律**：引用3条以上政策，每条包含发布机构+日期+全称+影响

#### 行业现状（第3页）
```
section标签: 行业现状
副标题: ——[数据驱动的增长判断]

[图表/数据展示]
数据来源：[权威机构]
```
**规律**：用柱状图/折线图展示2-3年数据趋势，标注来源

#### 痛点分析（第4页）
```
section标签: 痛点分析

[痛点名]（2-4字精炼命名）
[一句解释] + [量化数据]
```
**规律**：3-4个痛点，每个痛点用2-4字命名（如"数据孤岛"、"空域难估"），配量化数据+来源

#### 项目概述（第5页）
```
section标签: 项目概述

技术原理：[核心技术一句话概括]
技术突破：
  [指标1]：[具体数值]
  [指标2]：[具体数值]
  [指标3]：[具体数值]
  [指标4]：[具体数值]
前沿验证：
  [实地测试/试点成果]
技术应用：
  [场景1] [场景2] [场景3] [场景4]
```

#### 研发历程（第6页）
```
section标签: 研发历程

[YYYY.MM] [里程碑名称]
[一句话描述该阶段成果]
```
**规律**：5-10个里程碑，采用"时间+标题+描述"三要素，体现从0到1的完整过程

#### 核心技术（第7-N页，每个技术1-2页）
```
section标签: 核心技术
副标题: ——[技术名称]+[解决什么问题]

[技术名称]
[技术原理描述]
[关键性能指标，用数学符号量化]
衍生技术: [延伸应用]
前沿验证: [合作方验证+提升数据]
```
**规律**：
- 每个技术独立1-2页
- 第1页：技术原理+指标+验证
- 第2页：衍生技术+应用场景
- 所有指标必须量化（≤、≥、<、>、%）

#### 技术参数表（第N+1页）
```
section标签: 技术参数
副标题: ——多项数据重大突破，达到国内领先水平！

| 功能模块 | 核心指标 | 核心参数 |
|---------|---------|---------|
| [模块1] | [指标名] | [具体数值] |
| [模块2] | [指标名] | [具体数值] |
```

#### 应用案例（每个案例1页）
```
section标签: 应用案例

案例[X]：[企业名称]——[合作内容概括]
合作时间：[YYYY.MM]-[YYYY.MM]（[时长]）
合作企业：[企业全名]
企业资质：[一句话介绍企业背景]
合作内容：[我方提供什么 + 对方提供什么]
我方职责：[具体工作1] [具体工作2] [具体工作3]
合作成效：[量化成果1] [量化成果2] [量化成果3]
```
**规律**：每个案例必须包含5要素（企业+时间+内容+职责+成效），附证书/照片佐证

#### 竞品分析（第N+2页）
```
section标签: 竞品分析

| 维度 | 本项目 | 竞品A | 竞品B | 竞品C |
|------|-------|------|------|------|
| [指标1] | [领先数据] | [对比数据] | [对比数据] | [对比数据] |
```
**规律**：4-6个维度对比，本项目优势列用不同颜色标注

#### 商业模式（2-4页）
```
section标签: 商业模式

[服务类型名称]
服务客群：[目标客户]
服务内容：[具体服务描述]
解决痛点：[对应痛点分析页的痛点]
创造价值：[量化价值]
```
**规律**：
- 按客群/服务类型分类3-5类
- 每类必有四要素：服务内容+客群+痛点+价值
- 已签合同金额最有力

#### 行业认可（第N+3页）
```
section标签: 行业认可

[企业/机构名称]
[一句话介绍企业资质]
[认可评价原文/合作意向]
```

#### 媒体报道（第N+4页）
```
section标签: 媒体报道
副标题: ——权威媒体聚焦，多维视角展现[项目价值]

[媒体名](YYYY.MM)——报道主题
[媒体名](YYYY.MM)——报道主题
...
```

#### 个人成长三部曲（3页）

**第1页：价值筑基**
```
section标签: 个人成长·价值筑基

[YYYY年M月] [里程碑]
[一句话描述]
```
**规律**：时间线展示从入学到项目成果的完整成长轨迹

**第2页：能力跃迁**
```
section标签: 个人成长·能力跃迁
副标题: ——以项目为熔炉，实现五大核心能力跨越式提升

五大能力：专业技术能力 / 跨学科整合能力 / 团队领导能力 / 组织协调能力 / 商业转化能力
每项配[分数] + [一句描述]
```

**第3页：红色淬炼**
```
section标签: 个人成长·红色淬炼
副标题: ——将[项目]植入[国家战略]，让[精神]坚定[初心]

[YYYY.MM] 思想筑基·使命觉醒
[YYYY.MM] 实践淬火·[地区]践行
[YYYY.MM] 意志磨练·攻坚担当
```

#### 团队介绍（2-3页）
```
第1页：团队顾问
[姓名] [头衔] [提供什么支持]

第2页：团队负责人
[国家级项目] 技术负责人
[国家级竞赛] 国家级立项
[国家级中心] 学生负责人
[国家级大创] 项目主持人

第3页：团队成员
[竞赛名] 国家级[X]等奖
主要负责[具体工作]
```
**规律**：竞赛获奖必须写全称+等级，每人标注具体分工

#### 学科引领/带动教育（1-2页）
```
section标签: 学科引领
副标题: ——以项目为牵引，四维反哺"XXX"交叉育人生态

四个维度：课程融合 / 平台共建 / 科研育苗 / 标准反哺
每维配[具体行动] + [建设成果（量化）]
```

#### 社会价值（第N+5页）
```
section标签: 社会价值

[对仗排比句，如"破XXX，立XXX"]
[技术革新→产业升级→人才赋能→行业领航]
[结尾升华到国家战略，如"为东北全面振兴"]
```
**规律**：采用对仗/排比修辞，技术→产业→人才→标准递进，结尾升华

#### 盈利预估（第N+6页）
```
section标签: 盈利预估
副标题: ——合同签约锚定XX万基石收入，预计20XX年营收达XX万元

[已签合同金额]
[年份] [收入] [增长率]
年均复合增长率超XX%
```

#### 未来规划（第N+7页）
```
section标签: 未来规划

[YYYY] [阶段名]
[量化目标1] [量化目标2] [量化目标3]
```

#### 致谢页（最后1页）
```
[项目名之智/芯，驱动/擘画[行业]新篇]
[技术1]·[场景1]·[价值1]
共铸/共筑[复兴/强国]新章
```
**规律**：三行格式，第一行点题，第二行技术+场景排比，第三行升华到国家使命

### C. 评分维度对应的页面权重

中国国际大学生创新大赛评分标准：
- **个人成长（30分）**：对应"个人成长三部曲"+ "团队介绍" + "学科引领" + "带动教育"
- **项目创新（30分）**：对应"核心技术" + "技术参数" + "技术壁垒" + "衍生技术"
- **产业价值（25分）**：对应"商业模式" + "应用案例" + "应用场景" + "盈利预估" + "竞品分析"
- **团队协作（15分）**：对应"团队介绍" + "领导关怀" + "行业认可" + "媒体报道"

**页面分配建议**：
- 个人成长类：6-8页（占比约20%）
- 项目创新类：10-15页（占比约35%）
- 产业价值类：10-12页（占比约30%）
- 团队协作类：4-6页（占比约15%）

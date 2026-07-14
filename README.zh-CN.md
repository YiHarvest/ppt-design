# PPT Design

简体中文 | [English](README.md)

`ppt-design` 是一个 Codex 插件，它内部有三个 skill：

- `frontend-slides`：负责演示文稿工程、固定画布 HTML、导航、编辑能力与动效规范。
- `baoyu-design`：负责内容分析、商业叙事、视觉方向、设计系统与可编辑 PPTX 导出。
- `html-ppt`：提供可复用的主题、整套模板、页面布局、动画与演示运行时。

三者组合形成一套从内容分析、视觉设计、可编辑 HTML 制作到真正可编辑 PowerPoint 交付的完整工作流。

## 插件能做什么

- 制作投资人路演、产品发布、商业提案、内部汇报、会议演讲和培训课件。
- 分析本地项目、文档、现有网站或已有演示文稿，并整理成清晰的演讲叙事。
- 生成固定 1920 × 1080（16:9）画布的自包含 HTML 演示文稿。
- 加入键盘导航、页面转场、元素分步入场、减少动态效果支持，以及可选的演讲者备注。
- 建立统一的商业视觉体系，包括字体、色彩、间距、版式、图表和可复用组件。
- 使用 `html-ppt` 的主题、整套演示模板、页面布局、CSS 动画和 Canvas 特效快速完成成品。
- 优化已有 HTML 演示文稿，或把已有 PPT/PPTX 转换成带动画的网页演示。
- 将兼容的演示文稿导出为真正可编辑的 `.pptx`，文本、形状和受支持的图表保持原生可编辑，而不是整页截图。
- 按需输出 PDF、幻灯片 PNG 和预览页面等辅助文件。

## 三个 skill 如何协作

本插件使用时应在提示词中明确指定顺序：

1. 使用 `frontend-slides` 确定 HTML 演示文稿的工程结构、固定画布、导航、编辑行为与动效规范。
2. 使用 `baoyu-design` 分析源材料，建立演讲逻辑、商业视觉方向与设计系统。
3. 使用 `html-ppt` 的模板、主题、布局、动画和运行时完成演示文稿。
4. 预览并检查 HTML，确认后通过可编辑导出流程生成最终 `.pptx`。


## 安装

Codex 从已配置的 Marketplace 快照安装插件。将 `ppt-design` 注册到本地或 Git Marketplace 后，执行：

```bash
codex plugin add ppt-design@<marketplace-name>
```

如果使用非默认的本地 Marketplace，需要先配置 Marketplace 根目录：

```bash
codex plugin marketplace add /path/to/marketplace-root
codex plugin list
codex plugin add ppt-design@<marketplace-name>
```

`/home/yqy/Projects/PPTSkill/ppt-design` 是插件源码目录，本身不是 Marketplace 根目录。使用 `codex plugin add` 之前，需要先让某个 Marketplace 条目指向该插件源码。安装或更新插件后，请新建一个 Codex 对话，使三个内置 skill 被重新加载。

根据 Codex 界面的显示方式，安装后的 skill 可能带有插件命名空间，例如 `ppt-design:frontend-slides`；它们的内部名称仍然是 `frontend-slides`、`baoyu-design` 和 `html-ppt`。

## 快速开始

在新的 Codex 对话中提供源文件或项目路径，并给出清晰的演示需求。建议至少包括：

- 受众与演示目的；
- 演讲时长或目标页数；
- 演讲型低密度，还是阅读型高密度；
- 视觉风格与品牌约束；
- 需要交付的格式，例如可编辑 HTML、可编辑 PPTX、PDF 或 PNG；
- 是否需要演讲者备注或演讲者模式。

示例：

```text
请依次使用 frontend-slides、baoyu-design 和 html-ppt，基于
/path/to/project 制作一份投资人演示文稿。

受众：种子轮投资人
时长：10 分钟
形式：演讲型，约 10–12 页
风格：高级商业风、自信、简洁，以深色中性色为主，搭配一个强调色

先用 frontend-slides 确定可编辑的 1920×1080 HTML 工程、导航和动效规范；
再用 baoyu-design 分析项目，建立投资叙事和视觉系统；最后使用 html-ppt
的模板、主题、布局、动画与运行时完成演示文稿。

在制作完整演示文稿之前，先给我 3 个封面/风格预览方案。只使用项目中的
真实内容，检查每一页是否存在溢出或重叠，并交付：
1. 可编辑、带动画的 HTML 演示文稿；
2. 文本和形状真正可编辑的 PPTX；
3. 适合 10 分钟路演的演讲者备注。
```

## 通用提示词模板

```text
请把 frontend-slides、baoyu-design 和 html-ppt 作为一套演示文稿 pipeline 使用。

源材料：[项目路径、文档、URL、PPTX 或笔记]
受众：[受众]
目的：[路演、发布、汇报、培训、会议演讲等]
时长 / 页数：[数值]
内容密度：[演讲型 / 阅读型]
语言：[语言]
视觉方向：[风格、品牌、参考与限制]
交付格式：[可编辑 HTML / 可编辑 PPTX / PDF / PNG]

工作流：
1. 使用 frontend-slides 确定固定 1920×1080 的 HTML 工程结构、导航、
   编辑行为、无障碍与动效规范。
2. 使用 baoyu-design 分析源材料、组织叙事，并建立统一的商业设计系统。
3. 使用 html-ppt 完成主题、模板、布局、动画、演讲者功能与运行时。
4. 必要时先提供有代表性的视觉方向预览，确认后再制作完整演示文稿。
5. 检查内容准确性、溢出、重叠、动画、键盘导航和导出兼容性。
6. 交付要求的文件，并标明 PPTX 中无法保持可编辑的元素。

附加要求：
- 只使用真实源内容，不得虚构指标、客户或商业结论。
- 每页只表达一个清晰观点，内容过多时拆页，不要过度缩小字号。
- 所有资源应自包含，或复制到独立的输出目录中。
- 为可编辑 PPTX 优先使用原生文本、形状和图表。
- 仅在明确要求时加入演讲者备注与演讲者模式。
```

## 常见用法

### 从软件项目制作商业演示文稿

```text
使用 ppt-design pipeline 分析 /path/to/project，制作一份 12 页的商业产品
演示文稿。先交付可编辑 HTML，视觉方向确认后再导出可编辑 PPTX。
```

### 优化已有 HTML 演示文稿

```text
使用 frontend-slides、baoyu-design 和 html-ppt 优化
/path/to/deck.html。保留事实内容，强化商业叙事和视觉层级，加入有目的的
动效，并在桌面和手机视口下检查每一页，同时保持固定 16:9 画布不重排。
```

### 将 PowerPoint 转换为动画 HTML

```text
把 /path/to/source.pptx 转换成自包含、带动画的 HTML 演示文稿。保留原有
结构和资源，用 frontend-slides 与 html-ppt 重建导航和动效，再用
baoyu-design 修正视觉不一致。另附转换报告，列出无法支持的 PowerPoint 效果。
```

### 导出真正可编辑的 PPTX

```text
使用 baoyu-design 的可编辑导出流程，把已确认的 HTML 演示文稿导出为
可编辑 PPTX。文本、受支持的形状和图表必须保持原生，不要使用整页截图，
除非我明确同意栅格兜底。运行导出检查，并列出所有不可编辑元素。
```

## 演示控制

实际快捷键取决于所选模板和运行时。典型的成品可能支持：

| 按键 | 功能 |
| --- | --- |
| `←` / `→` 或 `Space` | 上一页 / 下一页 |
| `F` | 全屏 |
| `O` | 总览模式（如果支持） |
| `S` | 演讲者模式（如果已加入） |
| `N` | 演讲者备注抽屉（如果已加入） |
| `T` | 在兼容的 `html-ppt` 演示中切换主题 |
| `A` | 在兼容的 `html-ppt` 演示中切换动画 |
| `E` | 行内编辑（如果成品包含编辑器） |
| `Ctrl/Cmd + S` | 保存或下载修改（如果已启用编辑） |

最终 HTML 应说明它实际实现了哪些快捷键，不应默认认为每个模板都包含全部功能。

## 交付建议

- 在设计迭代阶段，将可编辑 HTML 作为可审查的内容与视觉基准。
- 把演示文件和复制的资源统一放入独立输出目录，不要散落在源项目中。
- 在 HTML 内容与视觉方向确认后，再进行可编辑 PPTX 导出。
- CSS 视觉通常比复杂浏览器特效更容易转换。Canvas、滤镜、混合模式、视频和高度定制动画，在 PowerPoint 中可能需要可编辑近似方案或栅格兜底。
- 浏览器动画与 PowerPoint 动画属于不同系统。可编辑导出器可以把受支持的 `data-anim` 约定转换成 PowerPoint 原生动画，但不能保证所有网页特效完全一致。

## 插件目录结构

```text
ppt-design/
├── .codex-plugin/
│   └── plugin.json
├── skills/
│   ├── frontend-slides/
│   ├── baoyu-design/
│   └── html-ppt/
├── component-lock.json
├── README.md
└── README.zh-CN.md
```

`component-lock.json` 记录每个组件的来源、版本、文件数量和目录树哈希。三个 skill 目录是各自原始组件的快照，不应随意合并或改名。

## 开发与验证

在插件根目录执行：

```bash
cd /home/yqy/Projects/PPTSkill/ppt-design

uv run --with pyyaml python \
  ~/.codex/skills/.system/plugin-creator/scripts/validate_plugin.py .

for skill in frontend-slides baoyu-design html-ppt; do
  uv run --with pyyaml python \
    ~/.codex/skills/.system/skill-creator/scripts/quick_validate.py \
    "skills/$skill"
done
```

更新组件时，应整体替换对应的 skill 快照，更新 `component-lock.json` 中的记录，重新验证插件和三个 skill，再从 Marketplace 重装插件，并在新的 Codex 对话中测试。

## 当前内置版本

精确快照信息见 [`component-lock.json`](component-lock.json)：

- `frontend-slides`：上游版本 `2.1.0`。
- `baoyu-design`：`2026-07-14` 本地快照。
- `html-ppt`：上游 `main` 分支的 `2026-07-14` 快照。

## 限制

- 这是一个插件集合，不是自动执行的 pipeline 引擎；需要完整工作流时，请在提示词中明确调用三个 skill。
- 成品质量取决于源材料的准确性和完整度。
- 可编辑 PPTX 导出适用于兼容的演示文稿结构，不能把任意网页自动转换为完全可编辑的 PowerPoint。
- 浏览器专属特效在导出 PPTX 时可能需要简化。
- 插件不会虚构商业数据、市场规模、增长指标、客户名称或引用。缺失证据时，请提供材料或明确授权检索。


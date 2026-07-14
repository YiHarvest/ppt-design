# PPT Design

[简体中文](README.zh-CN.md) | English

`ppt-design` is a Codex plugin that packages three complementary presentation skills without renaming or merging them:

- `frontend-slides` — presentation engineering, fixed-stage HTML, navigation, editing, and motion.
- `baoyu-design` — content analysis, storytelling, commercial visual direction, design systems, and editable PPTX export.
- `html-ppt` — reusable themes, slide templates, layouts, animations, and presentation runtime.

Together they provide a practical workflow for creating animated, editable HTML presentations and genuinely editable PowerPoint files.

## What the plugin can do

- Create investor pitches, product launches, business proposals, internal reports, conference talks, and training decks.
- Analyze a local project, documents, an existing website, or an existing deck and turn the material into a presentation narrative.
- Generate self-contained HTML decks with a fixed 1920 × 1080 (16:9) stage.
- Add keyboard navigation, transitions, staged element reveals, reduced-motion support, and optional presenter notes.
- Establish a coherent commercial visual system covering typography, color, spacing, layout, charts, and reusable components.
- Start from `html-ppt` themes, full-deck templates, page layouts, CSS animations, and Canvas effects.
- Improve an existing HTML deck or convert an existing PPT/PPTX into an animated web presentation.
- Export a compatible deck as a truly editable `.pptx`; text, shapes, and supported charts remain editable instead of becoming full-slide screenshots.
- Produce supporting outputs such as PDF, slide PNGs, and preview pages when requested.

## How the three skills work together

The plugin intentionally keeps the three skills independent. It does not add a fourth orchestration skill. State the desired sequence in your request:

1. Use `frontend-slides` to define the HTML deck architecture, fixed-stage rules, navigation, editing behavior, and animation conventions.
2. Use `baoyu-design` to analyze the source material and establish the story, commercial art direction, and design system.
3. Use `html-ppt` to implement the deck with its templates, themes, layouts, animations, and runtime.
4. Preview and verify the HTML, then use the editable export workflow to generate the final `.pptx`.

This separation makes each component replaceable and upgradable while preserving its original name and behavior.

## Installation

Codex installs plugins from a configured Marketplace snapshot. Once `ppt-design` has been registered in a local or Git Marketplace, install it with:

```bash
codex plugin add ppt-design@<marketplace-name>
```

For a non-default local Marketplace, configure the Marketplace root first:

```bash
codex plugin marketplace add /path/to/marketplace-root
codex plugin list
codex plugin add ppt-design@<marketplace-name>
```

The source checkout at `/home/yqy/Projects/PPTSkill/ppt-design` is the plugin directory, not a Marketplace root by itself. A Marketplace entry must point to this plugin source before `codex plugin add` can install it. After installing or updating the plugin, start a new Codex thread so the three bundled skills are loaded.

Depending on the Codex UI, installed skills may be displayed with the plugin namespace, for example `ppt-design:frontend-slides`. Their internal names remain `frontend-slides`, `baoyu-design`, and `html-ppt`.

## Quick start

In a new Codex thread, provide the source path or files and a clear presentation brief. A strong request includes:

- audience and presentation purpose;
- speaking duration or target slide count;
- speaker-led or reading-first density;
- visual tone or brand constraints;
- required outputs, such as editable HTML, editable PPTX, PDF, or PNG;
- whether speaker notes or presenter mode are required.

Example:

```text
Use frontend-slides, baoyu-design, and html-ppt in this order to create an
investor presentation for /path/to/project.

Audience: seed-stage investors
Duration: 10 minutes
Format: speaker-led, about 10–12 slides
Style: premium commercial, confident, concise, dark neutral with one accent color

First use frontend-slides to define the editable 1920×1080 HTML deck,
navigation, and motion conventions. Then use baoyu-design to analyze the
project, build the investor narrative, and establish the visual system.
Finally use html-ppt templates, themes, layouts, animations, and runtime to
finish the deck.

Show me 3 cover/style previews before completing the full deck. Use real
project content, verify every slide for overflow and overlap, and deliver:
1. an editable animated HTML presentation;
2. a genuinely editable PPTX with editable text and shapes;
3. speaker notes suitable for a 10-minute pitch.
```

## Reusable prompt template

```text
Use frontend-slides, baoyu-design, and html-ppt as one presentation pipeline.

Source: [project path, document, URL, PPTX, or notes]
Audience: [audience]
Purpose: [pitch, launch, report, training, conference, etc.]
Duration / slide count: [value]
Content density: [speaker-led / reading-first]
Language: [language]
Visual direction: [style, brand, references, constraints]
Required outputs: [editable HTML / editable PPTX / PDF / PNG]

Workflow:
1. Use frontend-slides to define a fixed 1920×1080 HTML architecture,
   navigation, editing behavior, accessibility, and motion conventions.
2. Use baoyu-design to analyze the source, structure the narrative, and create
   a coherent commercial design system.
3. Use html-ppt to implement the selected theme, templates, layouts,
   animations, presenter features, and runtime.
4. Preview representative visual directions when needed, then build the deck.
5. Verify content accuracy, overflow, overlap, animation behavior, keyboard
   navigation, and export compatibility.
6. Deliver the requested files and identify any elements that could not remain
   editable in PPTX.

Additional requirements:
- Use real source content; do not invent metrics, customers, or claims.
- Keep one clear message per slide and split overcrowded slides.
- Keep assets self-contained or copied into the output directory.
- Prefer native text, shapes, and charts for editable PPTX export.
- Include speaker notes and presenter mode only when requested.
```

## Common tasks

### Create a deck from a software project

```text
Use the ppt-design pipeline to inspect /path/to/project and create a 12-slide
commercial product deck. Deliver editable HTML first, get the visual direction
approved, then export an editable PPTX.
```

### Improve an existing HTML presentation

```text
Use frontend-slides, baoyu-design, and html-ppt to improve
/path/to/deck.html. Preserve its factual content, strengthen the business
story and visual hierarchy, add purposeful motion, and verify every slide at
desktop and phone viewport sizes without changing the 16:9 stage.
```

### Convert a PowerPoint deck to animated HTML

```text
Convert /path/to/source.pptx into a self-contained animated HTML deck. Preserve
the original structure and assets, rebuild navigation and motion with
frontend-slides and html-ppt, and use baoyu-design to correct visual
inconsistencies. Also provide a conversion report listing any unsupported
PowerPoint effects.
```

### Export a truly editable PPTX

```text
Export the approved HTML deck as an editable PPTX using the baoyu-design
editable export workflow. Keep text, supported shapes, and charts native.
Do not use full-slide screenshots unless I explicitly approve a raster fallback.
Run the export checks and report any non-editable elements.
```

## Presentation controls

Controls depend on the selected template and runtime. A typical generated deck supports:

| Key | Action |
| --- | --- |
| `←` / `→` or `Space` | Previous / next slide |
| `F` | Fullscreen |
| `O` | Overview, when supported |
| `S` | Presenter mode, when included |
| `N` | Speaker notes drawer, when included |
| `T` | Cycle themes in compatible `html-ppt` decks |
| `A` | Cycle animation styles in compatible `html-ppt` decks |
| `E` | Inline editing, when the generated deck includes the editor |
| `Ctrl/Cmd + S` | Save/download changes, when editing is enabled |

The final HTML should document the controls it actually implements; do not assume every template includes every feature.

## Output guidance

- Use editable HTML as the reviewable source of truth during design iteration.
- Keep presentation files and copied assets in a dedicated output directory rather than scattering them through the source project.
- Request editable PPTX export only after the HTML content and visual direction are approved.
- CSS-only visuals usually translate more reliably than complex browser effects. Canvas, filters, blend modes, videos, and highly custom animations may require an editable approximation or raster fallback in PowerPoint.
- Browser animation and PowerPoint animation systems are different. The editable exporter can preserve supported `data-anim` conventions as native PowerPoint builds, but exact parity is not guaranteed for every web effect.

## Plugin structure

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

`component-lock.json` records the source, source version, file count, and tree hash for each bundled component. The three skill directories are snapshots of their original components and should not be casually merged or renamed.

## Development and validation

Run these commands from the plugin root:

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

When updating a component, replace its complete skill snapshot, update its entry in `component-lock.json`, validate the plugin and all three skills, reinstall the plugin from its Marketplace, and test it in a new Codex thread.

## Bundled versions

See [`component-lock.json`](component-lock.json) for the exact snapshots currently included:

- `frontend-slides` — upstream version `2.1.0`.
- `baoyu-design` — local snapshot dated `2026-07-14`.
- `html-ppt` — upstream `main` snapshot dated `2026-07-14`.

## Limitations

- This plugin is a bundle, not an automatic pipeline engine; prompt the three skills explicitly when you want the full workflow.
- Quality depends on the accuracy and completeness of the source material.
- Editable PPTX export is available for compatible deck structures; arbitrary web pages are not automatically convertible into fully editable PowerPoint slides.
- Browser-specific effects may need simplification during PPTX export.
- The plugin does not fabricate business data, market size, traction, customer names, or citations. Provide or authorize research for any missing evidence.


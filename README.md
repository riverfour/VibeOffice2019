# VibeOffice Upgrade Implementation Plan: MS Office 2019 Modernization

## 1. Executive Summary & Strategy

**Goal:** Upgrade VibeOffice—a static, browser-based office suite—to provide a full **Microsoft Office 2019** visual interface, feature set, and interoperability standard across its three core applications:

1. **Quire** (Microsoft Word 2019 equivalent)
2. **Ledger** (Microsoft Excel 2019 equivalent)
3. **Lectern** (Microsoft PowerPoint 2019 equivalent)

### Core Principles
- **Lossless OOXML Preservation:** Maintain strict compliance with Microsoft Open XML (`.docx`, `.xlsx`, `.pptx`). Files opened and saved in VibeOffice must open in Microsoft Office 2019/2021/365 without repair prompts or lost attributes.
- **Vanilla Tech Stack:** Retain the zero-build-step, zero-framework, static HTML/CSS/JS architecture (`public/` root, `window.L` / `window.VO` global namespace).
- **Client-Side Privacy:** 100% browser-based execution with local file storage and zero back-end data transmission.

---

## 2. Current Architecture vs. MS Office 2019 Target State

| Dimension | Current VibeOffice State | Office 2019 Target State |
| :--- | :--- | :--- |
| **Default Interface** | Classic Office 2003 Toolbars & Menus (with optional `paper2016` Ribbon) | Modern Office 2019 Ribbon layout as default, featuring app-branded header themes |
| **App Branding** | Shared Blue theme variables | Distinct App Branding Accent Colors:<br>• **Quire (Word):** Word Blue (`#185abd`) <br>• **Ledger (Excel):** Excel Green (`#107c41`) <br>• **Lectern (PowerPoint):** PowerPoint Red-Orange (`#c43e1c`) |
| **Ribbon UI** | Simplified 2016 style | Office 2019 Ribbon tabs, Quick Access Toolbar (QAT) modern styling, contextual tab groups, and Tell Me search integration |
| **Word Processor (Quire)** | Flow layout engine, Markdown, HTML, `.docx` support | Focus Mode / Dark Canvas toggle, Learning Tools (line focus, column width adjustments), Math syntax builder, SVG icon graphics insertion |
| **Spreadsheet (Ledger)** | 519 formula functions, 2D charts, dynamic array engine | Office 2019 functions (`CONCAT`, `TEXTJOIN`, `IFS`, `SWITCH`, `MAXIFS`, `MINIFS`), Funnel charts, 2D Map chart rendering |
| **Presentations (Lectern)** | 15 layout designs, DrawingML renderers, animations | **Morph** transition renderer, modern 2019 slide transition effects, icon library, 4K slide export |

---

## 3. Phase 1: Visual Theme & UI Modernization (`office2019` Look)

### 3.1 CSS Theme System (`public/common/looks.css`)
Implement a new `[data-look="office2019"]` attribute selector in `public/common/looks.css` defining Office 2019 design tokens:

```css
/* Office 2019 Modern Theme */
[data-look="office2019"] {
  --theme-color: #185abd;
  --ui-font: "Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, "DejaVu Sans", sans-serif;
  --ui-size: 12px;
  --tray-radius: 0px;
  --ctl-radius: 2px;

  /* Office 2019 Canvas & Chrome */
  --surface: #ffffff;
  --canvas: #e6e6e6;
  --work: #f3f3f3;
  --chrome-line: #d1d1d1;
  --panel-border: #e1e1e1;

  /* Ribbon styling */
  --rb-bg: #f3f3f3;
  --rb-panel: #ffffff;
  --rb-line: #d1d1d1;
  --rb-tab-ink: #333333;
  --rb-tab-on-ink: var(--app-accent, #185abd);
  --rb-tab-hot: #eaeaea;
  --rb-file-bg: var(--app-accent, #185abd);
  --rb-file-ink: #ffffff;

  /* App Specific Accent Colors */
  --word-accent: #185abd;
  --excel-accent: #107c41;
  --powerpoint-accent: #c43e1c;
}
```

### 3.2 App-Specific Header Branding
Configure app-specific root selectors so each app automatically inherits its Office 2019 signature color:
- Quire: `html[data-app="quire"] { --app-accent: var(--word-accent); }`
- Ledger: `html[data-app="ledger"] { --app-accent: var(--excel-accent); }`
- Lectern: `html[data-app="lectern"] { --app-accent: var(--powerpoint-accent); }`

### 3.3 Look & Layout Selector Updates (`public/common/look.js`)
Update `public/common/look.js` to set `office2019` as the default theme look and ensure `data-layout` sets `ribbon`:

```javascript
// public/common/look.js
const l = VO.settings.get('vibeoffice.look') || 'office2019';
d.dataset.look = l;
d.dataset.layout = (l === 'paper2016' || l === 'office2019') ? 'ribbon' : 'bars';
```

---

## 4. Phase 2: Quire (Word Processor 2019 Parity)

### 4.1 Ribbon Redesign (`public/quire/js/ribbon.js`)
Structure the Quire Ribbon to match Word 2019:
- **File:** Open, Save, Save As, Export (PDF, HTML, MD, RTF), Options.
- **Home:** Clipboard, Font, Paragraph, Styles, Editing.
- **Insert:** Pages, Tables, Illustrations (Pictures, Shapes, Icons, SmartArt, Charts), Links, Header & Footer, Text (Text Box, LettersArt/WordArt, Drop Cap), Symbols (Equation, Symbol).
- **Draw:** Pen selection, Highlighter, Canvas drawing overlay.
- **Design:** Document Formatting, Page Color, Page Borders, Watermark.
- **Layout:** Page Setup (Margins, Orientation, Size, Columns), Indent & Spacing, Arrange.
- **References:** Table of Contents, Footnotes, Citations, Captions, Index.
- **Review:** Proofing (Spelling & Grammar), Word Count, Read Aloud, Translate, Comments, Tracking (Track Changes), Compare, Protect.
- **View:** Views (Print Layout, Web Layout, Focus Mode, Reading Mode), Show (Ruler, Gridlines, Navigation Pane), Zoom, Window.

### 4.2 Key Word 2019 Features
1. **Focus Mode & Dark Canvas:**
   * Implement a high-contrast dark page canvas mode in `public/quire/js/view.js` with inverted document color map (`html[data-dark-canvas="true"]`).
2. **Learning Tools / Immersive Reader:**
   * Text spacing adjustment toggle, column width regulator (Very Narrow / Narrow / Moderate / Wide), and line focus highlight overlay.
3. **LaTeX Math Syntax Input:**
   * Enhance `public/quire/js/math.js` to parse LaTeX math expressions (e.g., `\frac{a}{b}`) directly into Office Math Markup Language (OMML) / DrawingML math objects.

---

## 5. Phase 3: Ledger (Spreadsheet Engine 2019 Parity)

### 5.1 Ribbon Redesign (`public/ledger/js/ribbon.js`)
Structure the Ledger Ribbon to match Excel 2019:
- **Home:** Clipboard, Font, Alignment, Number, Styles (Conditional Formatting, Format as Table, Cell Styles), Cells (Insert, Delete, Format), Editing (AutoSum, Fill, Clear, Sort & Filter, Find & Select).
- **Insert:** Tables (PivotTable, Recommended PivotTables, Table), Illustrations, Charts (Recommended Charts, Column, Bar, Line, Pie, Hierarchy, Statistic, Scatter, Waterfall, Funnel, Combo), Sparklines, Filters (Slicer, Timeline), Links, Text, Symbols.
- **Page Layout:** Themes, Page Setup, Scale to Fit, Sheet Options, Arrange.
- **Formulas:** Insert Function, Function Library (AutoSum, Recently Used, Financial, Logical, Text, Date & Time, Lookup & Reference, Math & Trig, More Functions), Defined Names, Formula Auditing, Calculation.
- **Data:** Get & Transform Data, Queries & Connections, Sort & Filter, Data Tools (Text to Columns, Flash Fill, Remove Duplicates, Data Validation, Consolidate, What-If Analysis), Forecast, Outline.
- **Review:** Proofing, Accessibility, Language, Comments, Protect, Unprotect Sheet.
- **View:** Workbook Views (Normal, Page Break Preview, Page Layout), Show (Gridlines, Formula Bar, Headings), Zoom, Window, Macros.

### 5.2 Key Excel 2019 Features
1. **Formulas & Calculation Engine Verification:**
   * Ledger already supports 519 functions. Verify and optimize performance for key Office 2019 functions in `public/ledger/js/calc/`:
     * Logical: `IFS`, `SWITCH`
     * Text: `CONCAT`, `TEXTJOIN`
     * Statistical: `MAXIFS`, `MINIFS`
2. **Funnel & 2D Map Charts (`public/common/charts.js`):**
   * Extend `public/common/charts.js` SVG renderer to support:
     * **Funnel Chart:** Centered horizontal bar hierarchy with trapezoidal polygon connections.
     * **Map Chart (Choropleth):** SVG path region rendering for geographic data visualization.

---

## 6. Phase 4: Lectern (Presentation Editor 2019 Parity)

### 6.1 Ribbon Redesign (`public/lectern/js/ribbon.js`)
Structure the Lectern Ribbon to match PowerPoint 2019:
- **Home:** Clipboard, Slides (New Slide, Layout, Reset, Section), Font, Paragraph, Drawing, Editing.
- **Insert:** Slides, Tables, Images, Illustrations (Shapes, Icons, 3D Models, SmartArt, Chart), Add-ins, Links, Comments, Text, Symbols, Media (Video, Audio, Screen Recording).
- **Design:** Themes, Variants, Customize (Slide Size, Format Background), Design Ideas.
- **Transitions:** Transition to This Slide (None, Morph, Cut, Fade, Push, Wipe, Split, Reveal, Random Bars, Shape, Uncover, Cover, Flash), Effect Options, Timing (Sound, Duration, Apply To All, Advance Slide).
- **Animations:** Animation (Preview, Animation presets), Advanced Animation (Add Animation, Animation Pane, Trigger), Timing (Start, Duration, Delay, Reorder).
- **Slide Show:** Start Slide Show (From Beginning, From Current Slide, Present Online, Custom Slide Show), Set Up (Set Up Slide Show, Hide Slide, Rehearse Timings), Monitors (Use Presenter View).
- **Review:** Proofing, Accessibility, Language, Comments, Compare.
- **View:** Presentation Views (Normal, Outline View, Slide Sorter, Notes Page, Reading View), Master Views (Slide Master, Handout Master, Notes Master), Show (Ruler, Gridlines, Guides), Zoom, Color/Grayscale, Window.

### 6.2 Key PowerPoint 2019 Features
1. **Morph Transition Engine (`public/lectern/js/anim.js`):**
   * Implement object tracking across consecutive slides by shape name (`spPr` ID or `ph` type) and geometry interpolation (position `off`, size `ext`, rotation, fill color).
2. **Icon & Vector Graphic Library:**
   * Integrate modern SVG icon collection (`public/common/icons/office2019/`) with DrawingML path conversion.
3. **4K Video / High-Res Export:**
   * Add high-resolution canvas exporting capabilities (`3840x2160`) for PNG slide decks and MP4/WebM slide slideshow export.

---

## 7. Phase 5: Suite Integration & Start Center

### 7.1 Start Center Upgrade (`public/start.js` / `public/index.html`)
- Update the Start Center layout with Office 2019 visual aesthetics:
  - Clean sidebar navigation with app icons and accent badges.
  - Template preview grid with Office 2019 sample templates.
  - Quick Look switcher in the settings dialog enabling seamless toggling between **Office 2019**, **Paper 2016**, and **Classic 2003**.

### 7.2 Settings & Storage Key Sync
- Add configuration keys in `VO.settings`:
  - `'vibeoffice.look'` = `'office2019'`
  - `'vibeoffice.ribbonMode'` = `'full'` | `'simplified'`

---

## 8. Phase 6: Testing, OOXML Interoperability & QA

To comply with VibeOffice design rules and preservation contracts:

1. **OOXML Package Validation:**
   - Run `python3 tools/ooxml/package.py` to ensure all saved `.docx`, `.xlsx`, and `.pptx` files pass Open XML SDK validation rules without schema errors.
2. **Automated Visual Regression:**
   - Execute `node tools/shot.mjs quire/`
   - Execute `node tools/shot.mjs ledger/`
   - Execute `node tools/shot.mjs lectern/`
   - Verify zero console errors, zero external requests, and exact layout fidelity in Chromium CDP.
3. **Quick Test Suite:**
   - Run `node tools/quick.mjs` across modified files.
4. **Office 2019 / LibreOffice Interoperability Scan:**
   - Round-trip sample documents using `tests/ledger/roundtrip.js` and Lectern real-deck harnesses.
   - Verify that opened files display correctly in installed Microsoft Office 2019/365 and LibreOffice without repair dialogs.

---

## 9. Implementation Milestones

```
[ Milestone 1: Theme & CSS Framework ] ──► [ Milestone 2: Ribbon Specs ]
                                                   │
┌──────────────────────────────────────────────────┴──────────────────────────────────────────────────┐
│                                                                                                      │
▼                                                  ▼                                                  ▼
[ Milestone 3: Quire (Word) ]            [ Milestone 4: Ledger (Excel) ]            [ Milestone 5: Lectern (PPT) ]
│                                                  │                                                  │
└──────────────────────────────────────────────────┬──────────────────────────────────────────────────┘
                                                   │
                                                   ▼
                                  [ Milestone 6: QA & Office Validation ]
```

- [x] **Milestone 1:** Create `office2019` theme in `public/common/looks.css` & hook into `public/common/look.js`.
- [x] **Milestone 2:** Implement Office 2019 Ribbon specs in `public/common/ribbon.js` and app-specific `ribbon.js` files.
- [x] **Milestone 3:** Add Word 2019 Focus Mode and dark page canvas to Quire.
- [x] **Milestone 4:** Add Funnel chart renderer to `public/common/charts.js` and verify Excel 2019 formulas in Ledger.
- [x] **Milestone 5:** Implement Morph transition animation logic in Lectern (`public/lectern/js/slideshow.js`).
- [ ] **Milestone 6:** Execute full test suite (`quick.mjs`, `ooxml/package.py`, CDP screenshots) and verify OOXML compatibility.

---

## 10. Implementation Summary (Completed Phases 1-5)

### Phase 1: Visual Theme & UI Modernization ✅
- **`public/common/looks.css`**: Added `[data-look="office2019"]` theme with:
  - Flat, borderless ribbon design
  - App-branded accent colors (Word Blue `#185abd`, Excel Green `#107c41`, PowerPoint Red-Orange `#c43e1c`)
  - CSS variables for `--word-accent`, `--excel-accent`, `--powerpoint-accent`
  - Auto-switching `--app-accent` based on `html[data-app="..."]`
  - Focus Mode / Dark Canvas / Immersive Reader styles
  - Funnel and Map chart support
- **`public/common/look.js`**: Set `office2019` as default look with ribbon layout
- **`public/common/suite.js`**: Added `office2019` to theme list

### Phase 2: Quire (Word 2019 Parity) ✅
- **`public/quire/js/ribbon.js`**: Updated View tab with "Immersive" group containing:
  - Focus Mode toggle
  - Dark Canvas toggle
  - Column Width dropdown (Very Narrow, Narrow, Moderate, Wide)
  - Line Focus dropdown (None, One Line, Three Lines, Five Lines)
  - Text Spacing dropdown (Normal, Wide, Very Wide)
- **`public/quire/js/app.js`**: Implemented methods:
  - `focusMode()` - toggles focus mode with sidebar hiding
  - `darkCanvas()` - toggles dark page canvas
  - `columnWidthMenu()` / `setColumnWidth()` - column width control
  - `lineFocusMenu()` / `setLineFocus()` - line focus highlighting
  - `textSpacingMenu()` / `setTextSpacing()` - text spacing adjustment
- **`public/quire/js/commands.js`**: Added `focusMode` and `darkCanvas` commands
- **`public/quire/js/dialogs2.js`**: Enhanced Equation dialog with:
  - LaTeX syntax support (`\frac{}{}`, `\sqrt{}`, `_{}`, `^{}`, `\sum_{}^{}`, `\int_{}^{}`)
  - Symbol palette and LaTeX snippet buttons
  - Cambria Math font insertion

### Phase 3: Ledger (Excel 2019 Parity) ✅
- **`public/ledger/js/ribbon.js`**: Updated Insert → Charts group with dropdown for chart type selection
- **`public/common/charts.js`**: Added new chart types:
  - **Funnel Chart** (`funnel`): Centered horizontal bars with trapezoidal connections
  - **Map Chart** (`map`): Choropleth-style region coloring
- **`public/ledger/js/xchart.js`**: Updated Chart Wizard:
  - Added Funnel and Map to chart type gallery
  - Added SVG preview icons for both new chart types
  - Updated `groupOf()` to handle funnel/map chart types
- **Excel 2019 Functions** (already present in `fn-core.js`):
  - `IFS`, `SWITCH` (logical)
  - `CONCAT`, `TEXTJOIN` (text)
  - `MAXIFS`, `MINIFS` (statistical)

### Phase 4: Lectern (PowerPoint 2019 Parity) ✅
- **`public/lectern/js/ribbon.js`**: Updated Transitions tab:
  - Changed to dropdown menu with Morph transition icon
- **`public/lectern/js/slideshow.js`**: Implemented Morph transition:
  - Added `'morph'` to transitions list
  - Added `morphTransition()` function for shape interpolation
  - Added `transitionMenu()` for ribbon dropdown
  - Shape matching by `data-shape-id` attribute
  - Position, size, and opacity interpolation between slides
- **`public/common/icons.js`**: Added `morph` icon

### Phase 5: Suite Integration (Planned)
- Start Center theme switcher UI
- Settings synchronization for ribbon mode

---

## 11. Files Modified

| File | Phase | Description |
|------|-------|-------------|
| `public/common/looks.css` | 1 | Office 2019 theme, Focus Mode, Dark Canvas, Funnel/Map charts |
| `public/common/look.js` | 1 | Default look set to office2019 with ribbon layout |
| `public/common/suite.js` | 1 | Added office2019 to theme list |
| `public/common/icons.js` | 1, 4 | Added equation icon, morph transition icon |
| `public/common/charts.js` | 3, 4 | Funnel and Map chart renderers, chart type groups |
| `public/quire/js/ribbon.js` | 2 | View tab → Immersive group with Focus Mode, Dark Canvas |
| `public/quire/js/app.js` | 2 | Focus Mode, Dark Canvas, Immersive Reader methods |
| `public/quire/js/commands.js` | 2 | focusMode, darkCanvas commands |
| `public/quire/js/dialogs2.js` | 2 | Enhanced Equation dialog with LaTeX support |
| `public/ledger/js/ribbon.js` | 3 | Insert → Charts dropdown with type gallery |
| `public/ledger/js/xchart.js` | 3 | Chart Wizard with Funnel/Map types and preview icons |
| `public/lectern/js/ribbon.js` | 4 | Transitions tab → dropdown with Morph |
| `public/lectern/js/slideshow.js` | 4 | Morph transition implementation, transitionMenu() |

# BTA Column Chart

A Power BI custom visual — a column chart supporting **grouped, stacked, and 100% stacked** modes, a **hierarchical category axis** with per-level formatting, **data labels** and stack-total labels, customizable **tooltips** (including extra Tooltips-well measures), per-series/per-category color controls, a configurable legend, an analytics constant line, and a built-in scrollbar for large category sets.

Built on [Apache ECharts](https://echarts.apache.org/). The visual renders entirely within the Power BI sandbox and makes **no external network calls** — your data never leaves the report.

Part of the **BTA Charts** suite by **Beyond the Analytics**.

- **GUID:** `columnChartVisualA1B2C3D4E5F6A7B8`
- **Version:** 2.0.1.0
- **API version:** 5.11.0

## Build from source

Requires Node.js 18+ and npm.

```bash
npm install
npm run package    # → dist/columnChartVisualA1B2C3D4E5F6A7B8.2.0.1.0.pbiviz
```

Development server:

```bash
npm start          # serves the visual at https://localhost:8080 for the Power BI developer visual
```

Lint:

```bash
npm run lint
```

## Project layout

```
src/
  visual.ts        # IVisual entry — ECharts option building, host integration, tooltips, selection
  transformer.ts   # DataView → view model (pure, no ECharts/DOM)
  settings.ts      # FormattingSettingsModel (format-pane cards)
capabilities.json  # data roles + dataViewMappings + objects
pbiviz.json        # visual manifest
style/visual.less  # styles
assets/icon.png    # visual icon
```

## License

[MIT](./LICENSE) © Beyond the Analytics

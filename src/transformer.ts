"use strict";

/**
 * Transformer: Power BI categorical DataView → column-chart ViewModel.
 *
 * Pure data work — imports `powerbi-visuals-api` and nothing else (no echarts,
 * no DOM). Two shapes supported:
 *   - flat: `categorical.categories[]` + a single `values[]` measure.
 *   - grouped: legend role bound → `categorical.values.grouped()` returns one
 *     entry per legend group, each carrying the measure for that group.
 *
 * Up to four `category` columns can be bound for a hierarchical X-axis. Their
 * row-wise values are joined with " · " into a composite category key; the
 * outer-level value is exposed separately so the visual can draw a second
 * "parent group" axis underneath the leaf labels.
 */

import powerbi from "powerbi-visuals-api";

type DataView = powerbi.DataView;
type DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
type DataViewValueColumn = powerbi.DataViewValueColumn;
type PrimitiveValue = powerbi.PrimitiveValue;
type IVisualHost = powerbi.extensibility.visual.IVisualHost;
type ISelectionId = powerbi.visuals.ISelectionId;

export interface TooltipRow {
    displayName: string;
    value: PrimitiveValue;
    format?: string;
}

export interface ColumnSeries {
    /** Series label (legend entry, or measure name when no legend). */
    name: string;
    /** Default palette colour for the series. Per-cell overrides on
     *  `dataPoints` win when present. */
    color: string;
    /** Selection id for the whole legend group (legend-click cross-filter). */
    legendSelectionId: ISelectionId | null;
    /** Per-category point — `points[i]` aligns with `vm.categories[i]`. */
    points: ColumnPoint[];
}

export interface ColumnPoint {
    /** Aggregated raw value, or null when the row was missing. */
    value: number | null;
    /** Highlight value when a peer cross-highlighted this point. */
    highlight: number | null;
    /** Cross-filter selection id for the cell. */
    selectionId: ISelectionId | null;
    /** Per-cell colour override, if the user picked one via Data colors. */
    colorOverride?: string;
    tooltips: TooltipRow[];
}

export interface ColumnCategory {
    /** Composite label rendered on the leaf axis. */
    label: string;
    /** Per-level values (outer → leaf). When only one category is bound the
     *  array has a single entry equal to `label`. */
    levels: string[];
    /** Selection id for the whole category (cross-filters the report). */
    selectionId: ISelectionId | null;
}

export interface ColumnViewModel {
    categories: ColumnCategory[];
    series: ColumnSeries[];
    /** True when 2+ category fields are bound — drives the parent-group axis. */
    hierarchical: boolean;
    /** Display names of the bound category fields (outer → leaf). */
    categoryFieldNames: string[];
    /** Display name of the legend column when a legend is bound. */
    legendDisplayName: string;
    /** Measure display name (single measure: this is the series name when no legend). */
    valueDisplayName: string;
    /** PBI format string of the measure, used for tooltip / data-label formatting. */
    valueFormat?: string;
    /** True when at least one peer-highlighted point exists. */
    hasHighlights: boolean;
    isEmpty: boolean;
}

function emptyViewModel(): ColumnViewModel {
    return {
        categories: [],
        series: [],
        hierarchical: false,
        categoryFieldNames: [],
        legendDisplayName: "",
        valueDisplayName: "",
        valueFormat: undefined,
        hasHighlights: false,
        isEmpty: true,
    };
}

const DEFAULT_PALETTE = [
    "#118dff", "#12239e", "#e66c37", "#6b007b", "#e044a7",
    "#744ec2", "#d9b300", "#d64550", "#1aab40", "#5e5e5e",
];

function pickColor(host: IVisualHost, key: string, index: number): string {
    try {
        const palette = host?.colorPalette;
        if (palette?.getColor) return palette.getColor(key).value;
    } catch {
        // fall through to default
    }
    return DEFAULT_PALETTE[index % DEFAULT_PALETTE.length];
}

/** Read a per-point colour override written by the Data-colors card.
 *  PBI stores it under `<column>.objects[rowIdx].dataColors.fill.solid.color`. */
function readColorOverride(col: { objects?: { [k: number]: unknown } } | undefined, rowIdx: number): string | undefined {
    const o = col?.objects?.[rowIdx] as { dataColors?: { fill?: { solid?: { color?: string } } } } | undefined;
    return o?.dataColors?.fill?.solid?.color;
}

function toNumber(v: PrimitiveValue | undefined): number | null {
    if (v == null) return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
}

/** Degenerate grouped shape: the same field is bound to BOTH Axis and Legend, so
 *  PBI sends an empty `categories` array (the field appears only as the value
 *  grouping). Rather than blank out, collapse to a normal single-series column
 *  chart — one bar per group, at the group's total — like native PBI. Each bar
 *  is coloured distinctly so it still reads as "coloured by that field". */
function collapsedFromGroups(
    groups: Array<powerbi.DataViewValueColumnGroup>,
    valuesArr: powerbi.DataViewValueColumns,
    host: IVisualHost,
): ColumnViewModel {
    const categories: ColumnCategory[] = [];
    const points: ColumnPoint[] = [];
    let valueFormat: string | undefined;
    let valueDisplayName = "";

    const sumCol = (col: DataViewValueColumn | undefined): number =>
        (col?.values ?? []).reduce<number>((s, v) => s + (toNumber(v) ?? 0), 0);

    groups.forEach((group, gIdx) => {
        const cols = (group.values ?? []) as DataViewValueColumn[];
        // The grouped mapping carries both the value measure and any Tooltips-well
        // measures inside each group; split them by role.
        const measure = cols.find(c => c.source?.roles?.values) ?? cols[0];
        if (!measure) return;
        if (!valueFormat) valueFormat = measure.source?.format;
        if (!valueDisplayName) valueDisplayName = measure.source?.displayName ?? "";
        const name = group.name == null ? "(Blank)" : String(group.name);
        const total = sumCol(measure);

        // Tooltips-well measures → extra rows (NOT the value itself; the renderer
        // already shows the series value, so adding it here would duplicate it).
        const tooltips: TooltipRow[] = cols
            .filter(c => c.source?.roles?.tooltips)
            .map(c => ({
                displayName: c.source?.displayName ?? "",
                value: sumCol(c),
                format: c.source?.format,
            }));

        let selectionId: ISelectionId | null = null;
        try {
            selectionId = host.createSelectionIdBuilder()
                .withSeries(valuesArr, group)
                .createSelectionId();
        } catch {
            selectionId = null;
        }

        categories.push({ label: name, levels: [name], selectionId });
        points.push({
            value: total,
            highlight: null,
            selectionId,
            colorOverride: pickColor(host, name, gIdx),
            tooltips,
        });
    });

    if (categories.length === 0) return emptyViewModel();

    return {
        categories,
        series: [{
            name: valueDisplayName,
            color: pickColor(host, valueDisplayName || "series", 0),
            legendSelectionId: null,
            points,
        }],
        hierarchical: false,
        categoryFieldNames: [valueDisplayName],
        legendDisplayName: "",
        valueDisplayName,
        valueFormat,
        hasHighlights: false,
        isEmpty: false,
    };
}

export function transform(
    dataView: DataView | undefined,
    host: IVisualHost,
): ColumnViewModel {
    const cat = dataView?.categorical;
    if (!cat || !cat.values?.length) return emptyViewModel();

    // Detect grouped (legend bound) shape — group identity is present on the
    // first group when `values.group.by` was used.
    const valuesArr = cat.values;
    const valuesArrAny = valuesArr as unknown as {
        grouped?: () => Array<powerbi.DataViewValueColumnGroup>;
        source?: { displayName?: string };
    };
    const groups = typeof valuesArrAny.grouped === "function" ? valuesArrAny.grouped() : null;

    // No category axis: either the same field is bound to Axis + Legend (PBI
    // drops `categories` and surfaces the field only as the grouping), or only a
    // legend was bound. Collapse the groups to a single-series axis instead of
    // blanking — see collapsedFromGroups.
    if (!cat.categories?.length) {
        return groups && groups.length > 0
            ? collapsedFromGroups(groups, valuesArr, host)
            : emptyViewModel();
    }

    const hasLegend = !!(groups && groups.length > 0 && groups[0]?.identity);

    // Resolve category columns by role — supports a hierarchical X-axis.
    const allCategoryCols = (cat.categories as DataViewCategoryColumn[])
        .filter(c => c.source?.roles?.category);
    if (allCategoryCols.length === 0) allCategoryCols.push(cat.categories[0]);
    const primaryCategoryCol = allCategoryCols[0];
    const hierarchical = allCategoryCols.length > 1;
    const categoryFieldNames = allCategoryCols.map(c => c.source?.displayName ?? "");

    // ── Categories ─────────────────────────────────────────────────────────
    const rowCount = primaryCategoryCol.values?.length ?? 0;
    const categories: ColumnCategory[] = [];
    const seenCategoryKeys = new Set<string>();
    const rowIndexByCategory = new Map<string, number>();

    for (let i = 0; i < rowCount; i++) {
        const levels = allCategoryCols.map(c => {
            const raw = (c.values ?? [])[i];
            return raw == null ? "(Blank)" : String(raw);
        });
        const label = hierarchical ? levels.join(" · ") : levels[0];
        if (seenCategoryKeys.has(label)) continue;
        seenCategoryKeys.add(label);
        rowIndexByCategory.set(label, i);

        let selectionId: ISelectionId | null = null;
        try {
            // ONE .withCategory() call only — chaining multiple .withCategory
            // calls is a known PBI bug that silently breaks cross-filtering.
            selectionId = host.createSelectionIdBuilder()
                .withCategory(primaryCategoryCol, i)
                .createSelectionId();
        } catch {
            selectionId = null;
        }
        categories.push({ label, levels, selectionId });
    }

    if (categories.length === 0) return emptyViewModel();

    // ── Tooltip columns (collect once for the whole point) ─────────────────
    const tooltipCats = (cat.categories ?? []).filter(c => c.source?.roles?.tooltips);
    const tooltipVals = (cat.values ?? []).filter((v: DataViewValueColumn) => v.source?.roles?.tooltips);

    const tooltipsForRow = (rowIdx: number): TooltipRow[] => {
        const out: TooltipRow[] = [];
        for (const tc of tooltipCats) {
            out.push({
                displayName: tc.source?.displayName ?? "",
                value: (tc.values ?? [])[rowIdx] ?? null,
                format: tc.source?.format,
            });
        }
        for (const tv of tooltipVals) {
            out.push({
                displayName: tv.source?.displayName ?? "",
                value: (tv.values ?? [])[rowIdx] ?? null,
                format: tv.source?.format,
            });
        }
        return out;
    };

    // ── Series build-out ───────────────────────────────────────────────────
    const series: ColumnSeries[] = [];
    let hasHighlights = false;
    let valueFormat: string | undefined;
    let valueDisplayName = "";
    let legendDisplayName = "";

    if (hasLegend && groups) {
        // Grouped: one series per legend group, each with one measure column.
        legendDisplayName = (valuesArrAny.source?.displayName)
            ?? groups[0]?.values?.[0]?.source?.displayName ?? "";
        groups.forEach((group, gIdx) => {
            const measure = group.values?.[0] as DataViewValueColumn | undefined;
            if (!measure) return;
            if (!valueFormat) valueFormat = measure.source?.format;
            if (!valueDisplayName) valueDisplayName = measure.source?.displayName ?? "";
            const seriesName = group.name == null ? "(Blank)" : String(group.name);
            const seriesKey = seriesName + "|" + String(gIdx);
            const baseColor = pickColor(host, seriesKey, gIdx);

            let legendSelectionId: ISelectionId | null = null;
            try {
                legendSelectionId = host.createSelectionIdBuilder()
                    .withSeries(valuesArr, group)
                    .createSelectionId();
            } catch {
                legendSelectionId = null;
            }

            const points: ColumnPoint[] = [];
            for (let i = 0; i < categories.length; i++) {
                const rowIdx = rowIndexByCategory.get(categories[i].label) ?? -1;
                const v = rowIdx < 0 ? null : toNumber(measure.values?.[rowIdx]);
                const hlRaw = rowIdx < 0 ? null : measure.highlights?.[rowIdx];
                const hl = hlRaw == null ? null : toNumber(hlRaw);
                if (hl != null) hasHighlights = true;

                let selectionId: ISelectionId | null = null;
                try {
                    selectionId = host.createSelectionIdBuilder()
                        .withCategory(primaryCategoryCol, rowIdx)
                        .withSeries(valuesArr, group as unknown as powerbi.DataViewValueColumnGroup)
                        .createSelectionId();
                } catch {
                    selectionId = null;
                }

                points.push({
                    value: v,
                    highlight: hl,
                    selectionId,
                    colorOverride: readColorOverride(primaryCategoryCol as unknown as { objects?: { [k: number]: unknown } }, rowIdx),
                    tooltips: rowIdx < 0 ? [] : tooltipsForRow(rowIdx),
                });
            }

            series.push({ name: seriesName, color: baseColor, legendSelectionId, points });
        });
    } else {
        // Flat (no legend) — one series per bound measure.
        const measures = (cat.values as DataViewValueColumn[])
            .filter(v => v.source?.roles?.values);
        const effectiveMeasures = measures.length > 0 ? measures : [cat.values[0] as DataViewValueColumn];

        effectiveMeasures.forEach((measure, mIdx) => {
            if (!valueFormat) valueFormat = measure.source?.format;
            const measureName = measure.source?.displayName ?? `Measure ${mIdx + 1}`;
            if (!valueDisplayName) valueDisplayName = measureName;
            const baseColor = pickColor(host, measureName, mIdx);

            const points: ColumnPoint[] = [];
            for (let i = 0; i < categories.length; i++) {
                const rowIdx = rowIndexByCategory.get(categories[i].label) ?? -1;
                const v = rowIdx < 0 ? null : toNumber(measure.values?.[rowIdx]);
                const hlRaw = rowIdx < 0 ? null : measure.highlights?.[rowIdx];
                const hl = hlRaw == null ? null : toNumber(hlRaw);
                if (hl != null) hasHighlights = true;

                let selectionId: ISelectionId | null = null;
                try {
                    selectionId = host.createSelectionIdBuilder()
                        .withCategory(primaryCategoryCol, rowIdx)
                        .createSelectionId();
                } catch {
                    selectionId = null;
                }

                points.push({
                    value: v,
                    highlight: hl,
                    selectionId,
                    colorOverride: readColorOverride(primaryCategoryCol as unknown as { objects?: { [k: number]: unknown } }, rowIdx),
                    tooltips: rowIdx < 0 ? [] : tooltipsForRow(rowIdx),
                });
            }

            series.push({ name: measureName, color: baseColor, legendSelectionId: null, points });
        });
    }

    if (series.length === 0) return emptyViewModel();

    return {
        categories,
        series,
        hierarchical,
        categoryFieldNames,
        legendDisplayName,
        valueDisplayName,
        valueFormat,
        hasHighlights,
        isEmpty: false,
    };
}

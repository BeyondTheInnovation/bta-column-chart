"use strict";

/**
 * Power BI column-chart visual built on Apache ECharts.
 *
 *   - one `bar` series per legend group (or per measure when no legend)
 *   - `xAxis.type: "category"` (leaf labels) + an optional parallel `xAxis` for
 *     the outer parent-group labels when a hierarchical X-axis is bound
 *   - `yAxis.type: "value"`
 *   - `barMode = "grouped" | "stacked" | "percent"` chooses between
 *     side-by-side bars, `stack: "total"`, and a percent-normalised stack
 *   - mouse-wheel-pannable `slider` dataZoom when categories overflow
 */

import powerbi from "powerbi-visuals-api";
import "./../style/visual.less";

import * as echarts from "echarts";
import type { EChartsType, EChartsOption } from "echarts";

import { FormattingSettingsService, formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ITooltipService = powerbi.extensibility.ITooltipService;
import ISelectionId = powerbi.visuals.ISelectionId;
import DataView = powerbi.DataView;

import { EChartsColumnChartFormattingSettings } from "./settings";
import { transform, type ColumnViewModel, type ColumnSeries } from "./transformer";

// ─── Helpers ────────────────────────────────────────────────────────────────

function abbreviate(v: number): string {
    const abs = Math.abs(v);
    if (abs >= 1e9) return (v / 1e9).toFixed(abs >= 1e10 ? 0 : 1) + "B";
    if (abs >= 1e6) return (v / 1e6).toFixed(abs >= 1e7 ? 0 : 1) + "M";
    if (abs >= 1e3) return (v / 1e3).toFixed(abs >= 1e4 ? 0 : 1) + "k";
    if (Number.isInteger(v)) return v.toLocaleString();
    return v.toFixed(2);
}

function formatValue(value: number | null | undefined, format: string | undefined): string {
    if (value == null || !Number.isFinite(value)) return "—";
    if (format && format.includes("%")) {
        const m = format.match(/0(?:\.(0+))?%/);
        const decimals = m?.[1]?.length ?? 0;
        return (value * 100).toFixed(decimals) + "%";
    }
    if (format && format.startsWith("$")) {
        return Math.abs(value) >= 10000 ? "$" + abbreviate(value)
            : "$" + Math.round(value).toLocaleString();
    }
    return Math.abs(value) >= 10000 ? abbreviate(value)
        : Number.isInteger(value) ? value.toLocaleString() : value.toFixed(2);
}

/** Format a number honouring a display-unit + decimal-places choice (used by
 *  the Y-axis and data/total labels). "auto" falls back to `abbreviate`. */
function fmtWithUnit(value: number | null | undefined, unit: string, decimals: number): string {
    if (value == null || !Number.isFinite(value)) return "—";
    const d = Math.max(0, Math.min(6, Math.trunc(decimals) || 0));
    const fix = (x: number, min: number) => x.toFixed(Math.max(min, d));
    switch (unit) {
        case "none": return d > 0 ? value.toFixed(d) : Math.round(value).toLocaleString();
        case "thousands": return fix(value / 1e3, 1) + "K";
        case "millions": return fix(value / 1e6, 1) + "M";
        case "billions": return fix(value / 1e9, 1) + "B";
        case "auto":
        default: return d > 0
            ? (Math.abs(value) >= 1e4 ? abbreviate(value) : value.toFixed(d))
            : abbreviate(value);
    }
}

function hexToRgba(hex: string, alpha: number): string {
    const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
    if (!m) return hex;
    const n = parseInt(m[1], 16);
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${alpha})`;
}

function escapeHtml(s: string): string {
    return String(s).replace(/[&<>"']/g, c => ({
        "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    }[c] as string));
}

function idKey(id: ISelectionId | null): string {
    if (!id) return "";
    const i = id as unknown as { getKey?: () => string };
    try { return i.getKey?.() ?? ""; } catch { return ""; }
}

// Native ECharts zoom+pan slider (modeled on the "mix-zoom-on-value" example):
// drag the end handles to ZOOM the window, drag the middle to PAN, and the
// track shows a data-shadow silhouette of the full series. It is paired with an
// `inside` dataZoom (wheel-zoom + drag-pan inside the plot). `filterMode:filter`
// drops off-window categories (clean edges, no half-bars) and lets the value
// axis rescale to the visible window. Animation stays off (gated at the root
// while scrollable) so neither zoom nor pan ever tweens the bars.
const SLIDER_STYLE = {
    type: "slider" as const,
    show: true,
    realtime: true,
    showDetail: false,
    showDataShadow: true,
    brushSelect: false,
    backgroundColor: "rgba(0,0,0,0.03)",
    fillerColor: "rgba(96,98,98,0.14)",
    borderColor: "transparent",
    dataBackground: {
        lineStyle: { color: "rgba(96,98,98,0.5)", width: 1 },
        areaStyle: { color: "rgba(96,98,98,0.18)" },
    },
    selectedDataBackground: {
        lineStyle: { color: "rgba(59,130,246,0.85)", width: 1 },
        areaStyle: { color: "rgba(59,130,246,0.3)" },
    },
    handleSize: "120%",
    handleStyle: { color: "#ffffff", borderColor: "#9ca3af", borderWidth: 1 },
    moveHandleSize: 5,
    moveHandleStyle: { color: "#bcbfc2" },
    emphasis: { handleStyle: { borderColor: "#3b82f6" }, moveHandleStyle: { color: "#3b82f6" } },
    throttle: 0,
    filterMode: "filter" as const,
    animation: false,
    animationDurationUpdate: 0,
};

// ─── Visual ─────────────────────────────────────────────────────────────────

export class Visual implements IVisual {
    private readonly root: HTMLDivElement;
    private readonly chartEl: HTMLDivElement;
    private readonly emptyNode: HTMLDivElement;
    private chart: EChartsType | null = null;

    private readonly host: IVisualHost;
    private readonly tooltipService: ITooltipService | null;
    private readonly selectionManager: ISelectionManager;
    private readonly formattingSettingsService: FormattingSettingsService;
    private formattingSettings: EChartsColumnChartFormattingSettings;
    private viewModel: ColumnViewModel | null = null;
    private viewportW = 0;
    private viewportH = 0;
    private allowInteractions = true;
    private chartEventsBound = false;

    /** Currently-selected ids — local mirror of selectionManager + the
     *  per-render highlight overlay. */
    private localSelectedKeys: Set<string> = new Set();
    private scrollEnabled = false;
    private totalCategories = 0;
    /** Live zoom window as start/end PERCENTAGES (0–100), carried across
     *  update() so a peer cross-filter / resize / format change keeps the
     *  user's current zoom + pan instead of snapping back. -1 = uninitialised
     *  (use the default initial window). */
    private scrollStartPct = -1;
    private scrollEndPct = -1;
    /** True after a data-change rebuild on a SCROLLABLE chart: the init "grow
     *  up" animation is playing, and once it finishes we disable animation so
     *  the native dataZoom zoom/pan never tweens (ECharts' realtime roam uses a
     *  200ms tween whenever isAnimationEnabled() is true, bypassing
     *  animationDurationUpdate:0). Non-scrollable charts keep animation on. */
    private pendingIntroDisable = false;
    /** Timer that disables animation ~after the intro completes (scrollable
     *  charts only) so the native dataZoom never tweens. */
    private introTimer: ReturnType<typeof setTimeout> | null = null;
    /** Identity of the last rendered dataset; a change re-enables the enter
     *  animation, an unchanged value (resize / format / selection rebuild)
     *  suppresses the otherwise-replayed 600ms intro. */
    private lastDataSig = "";
    /** Monotonic guard so out-of-order select().then resolutions can't clobber
     *  a newer click. */
    private clickSeq = 0;
    /** Re-entrancy guard while reverting ECharts' native legend toggle. */
    private legendSyncing = false;
    /** ctrl/⌘ captured on mousedown — legendselectchanged carries no MouseEvent. */
    private lastPointerMulti = false;
    /** Keyboard focus cursor across (series, category). */
    private focusedCat = -1;
    private focusedSeries = 0;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.tooltipService = this.host.tooltipService ?? null;
        this.selectionManager = this.host.createSelectionManager();
        const caps = (this.host as IVisualHost & { hostCapabilities?: { allowInteractions?: boolean } })
            .hostCapabilities;
        this.allowInteractions = caps?.allowInteractions !== false;
        this.formattingSettingsService = new FormattingSettingsService();
        this.formattingSettings = new EChartsColumnChartFormattingSettings();

        this.root = document.createElement("div");
        this.root.className = "echarts-column-chart-visual";
        this.chartEl = document.createElement("div");
        this.chartEl.className = "echarts-column-chart-visual__chart";
        // Keyboard focus (capabilities.supportsKeyboardFocus): arrow keys move
        // a focus cursor across bars, Enter/Space selects, Esc clears.
        this.chartEl.tabIndex = 0;
        this.chartEl.setAttribute("role", "application");
        this.chartEl.setAttribute("aria-label", "Column chart");
        this.chartEl.addEventListener("keydown", (e) => this.handleKeyDown(e));
        // zrender's `mousewheel` is a passive listener — preventDefault there
        // is a no-op, so own the wheel here too to stop the page scrolling.
        this.chartEl.addEventListener("wheel", (e) => {
            if (this.scrollEnabled) e.preventDefault();
        }, { passive: false });

        this.emptyNode = document.createElement("div");
        this.emptyNode.className = "echarts-column-chart-visual__empty";
        this.emptyNode.style.display = "none";
        this.buildEmptyState();

        this.root.appendChild(this.chartEl);
        this.root.appendChild(this.emptyNode);
        options.element.appendChild(this.root);

        this.selectionManager.registerOnSelectCallback?.((ids) => {
            this.localSelectedKeys = new Set((ids as ISelectionId[]).map(idKey));
            this.applySelectionDim();
        });
    }

    private buildEmptyState(): void {
        const skeleton = document.createElement("div");
        skeleton.className = "echarts-column-chart-visual__skeleton";
        skeleton.setAttribute("aria-hidden", "true");
        const heights = [45, 70, 55, 90, 60, 80, 50];
        for (const h of heights) {
            const bar = document.createElement("div");
            bar.className = "echarts-column-chart-visual__skeleton-bar";
            bar.style.height = `${h}%`;
            skeleton.appendChild(bar);
        }
        this.emptyNode.appendChild(skeleton);
    }

    public update(options: VisualUpdateOptions): void {
        const eventService = this.host.eventService;
        eventService?.renderingStarted?.(options);

        try {
            const dataView: DataView | undefined = options.dataViews?.[0];
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                EChartsColumnChartFormattingSettings,
                dataView ?? ({ metadata: {} } as DataView),
            );
            this.viewModel = transform(dataView, this.host);

            // Surface per-cell override colours back into the view-model.
            this.applyDataColorOverrides(this.viewModel);

            const w = Math.max(1, options.viewport?.width ?? 0);
            const h = Math.max(1, options.viewport?.height ?? 0);
            this.viewportW = w;
            this.viewportH = h;
            this.root.style.width = `${w}px`;
            this.root.style.height = `${h}px`;
            this.chartEl.style.width = `${w}px`;
            this.chartEl.style.height = `${h}px`;

            if (!this.viewModel || this.viewModel.isEmpty) {
                this.emptyNode.style.display = "flex";
                this.chartEl.style.visibility = "hidden";
                this.chart?.clear();
                eventService?.renderingFinished?.(options);
                return;
            }
            this.emptyNode.style.display = "none";
            this.chartEl.style.visibility = "visible";

            if (!this.chart) {
                // Canvas renderer over SVG for column-chart: each wheel-tick
                // re-renders every visible bar; SVG creates a DOM node per
                // bar which compounds at 100+ categories. Canvas keeps the
                // scroll latency under one frame.
                this.chart = echarts.init(this.chartEl, undefined, { renderer: "canvas" });
            } else {
                this.chart.resize();
            }

            // Carry the live scroll window across the rebuild. PBI fires
            // update() for resize / format tweaks / peer cross-filter; clear()
            // (needed to reset stale dataZoom state) would otherwise snap the
            // user back to the first categories. captureScroll() syncs
            // this.scrollStart from the live dataZoom (covers slider drags too),
            // and buildOption re-opens the window there.
            this.captureScroll();
            const option = this.buildOption(this.viewModel);
            // clear() before setOption — `notMerge: true` alone does not reset
            // dataZoom window state, so scroll position would leak across updates.
            this.chart.clear();
            this.chart.setOption(option, { notMerge: true });

            if (!this.chartEventsBound) {
                this.bindChartEvents();
                this.chartEventsBound = true;
            }
            // NOTE: do NOT call applySelectionDim() here — buildOption already
            // bakes the correct per-bar selection/highlight opacity into the
            // series data. A second setOption in the same tick would re-emit the
            // bars with animationDurationUpdate:0 and CLOBBER the just-started
            // init "grow up" animation. applySelectionDim() is only for live
            // selection changes (the click / cross-filter callbacks).

            // Scrollable chart + data changed → the intro "grow up" animation
            // is now playing. Once it finishes, disable animation so the native
            // dataZoom zoom/pan never tweens (ECharts' realtime roam uses a
            // ~200ms tween whenever isAnimationEnabled() is true). A fixed timer
            // is more reliable than the `finished` event, which can fire before
            // the intro. Non-scrollable charts keep animation (no scroll).
            if (this.introTimer) { clearTimeout(this.introTimer); this.introTimer = null; }
            if (this.pendingIntroDisable) {
                const chart = this.chart;
                this.introTimer = setTimeout(() => {
                    this.introTimer = null;
                    if (!chart) return;
                    const n = ((chart.getOption() as { series?: unknown[] }).series ?? []).length;
                    chart.setOption({
                        animation: false,
                        series: Array.from({ length: n }, () => ({ animation: false })),
                    }, false);
                }, 1150);
            }

            eventService?.renderingFinished?.(options);
        } catch (e) {
            console.error("[echarts-column-chart] update failed", e);
            eventService?.renderingFailed?.(options, e instanceof Error ? e.message : String(e));
        }
    }

    public destroy(): void {
        if (this.introTimer) { clearTimeout(this.introTimer); this.introTimer = null; }
        try { this.chart?.dispose(); } catch { /* ignore */ }
        this.chart = null;
    }

    // ─── Settings helpers ────────────────────────────────────────────────

    private enumValue(slot: { value?: unknown } | undefined, fallback: string): string {
        const v = slot?.value;
        if (v == null) return fallback;
        if (typeof v === "string") return v;
        if (typeof v === "object" && typeof (v as { value?: unknown }).value === "string") {
            return (v as { value: string }).value;
        }
        return fallback;
    }

    private colorOf(slot: { value?: unknown } | undefined, fallback: string): string {
        return (slot?.value as { value?: string } | undefined)?.value ?? fallback;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        try {
            const fs = this.formattingSettings;
            const vm = this.viewModel;

            // ── Dynamic "Data colors" card — one ColorPicker per distinct
            //    series (legend mode) or per category (no legend).
            const distinct: Array<{ name: string; color: string; selectionId: ISelectionId | null }> = [];
            if (vm && !vm.isEmpty) {
                if (vm.series.length > 1) {
                    for (const s of vm.series) {
                        distinct.push({ name: s.name, color: s.color, selectionId: s.legendSelectionId });
                    }
                } else if (vm.series.length === 1) {
                    const s = vm.series[0];
                    for (let i = 0; i < vm.categories.length; i++) {
                        const pt = s.points[i];
                        distinct.push({
                            name: vm.categories[i].label,
                            color: pt?.colorOverride ?? s.color,
                            selectionId: vm.categories[i].selectionId,
                        });
                    }
                }
            }
            if (distinct.length === 0) {
                (fs.dataColors as { visible?: boolean }).visible = false;
                fs.dataColors.slices = [];
            } else {
                (fs.dataColors as { visible?: boolean }).visible = true;
                const CONSTANT_OR_RULE =
                    (powerbi as unknown as { VisualEnumerationInstanceKinds?: { ConstantOrRule?: number } })
                        ?.VisualEnumerationInstanceKinds?.ConstantOrRule ?? 3;
                fs.dataColors.slices = distinct.map(d => new formattingSettings.ColorPicker({
                    name: "fill",
                    displayName: d.name,
                    value: { value: d.color },
                    selector: (d.selectionId as unknown as { getSelector?: () => unknown })?.getSelector?.() as never,
                    instanceKind: CONSTANT_OR_RULE,
                }));
            }

            // ── Conditional visibility (collapse irrelevant slices).
            const barMode = this.enumValue(fs.columns.barMode, "grouped");
            (fs.columns.seriesGap as { visible?: boolean }).visible = barMode === "grouped";
            // Parent-level X-axis controls appear only when formatting separately.
            const parentSep = !!fs.xAxis.parentSeparate.value;
            for (const s of [fs.xAxis.parentFontFamily, fs.xAxis.parentFontSize, fs.xAxis.parentBold, fs.xAxis.parentColor]) {
                (s as { visible?: boolean }).visible = parentSep;
            }
        } catch (e) {
            console.error("[echarts-column-chart] getFormattingModel", e);
        }
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // ─── Data colour overrides ───────────────────────────────────────────

    /** Walk per-cell `colorOverride` from the transformer into the actual
     *  per-point `itemStyle` used by `buildOption`. The transformer reads
     *  the raw `objects[i].dataColors.fill.solid.color` value off the
     *  category column; we keep that in place but also let an entire series
     *  inherit its override when the legend group has one. */
    private applyDataColorOverrides(vm: ColumnViewModel | null): void {
        if (!vm) return;
        for (const s of vm.series) {
            const firstOverride = s.points.find(p => p.colorOverride)?.colorOverride;
            if (s.points.every(p => p.colorOverride === firstOverride) && firstOverride) {
                s.color = firstOverride;
            }
        }
    }

    // ─── ECharts option building ─────────────────────────────────────────

    private buildOption(vm: ColumnViewModel): EChartsOption {
        const fs = this.formattingSettings;
        const palette = this.host.colorPalette as {
            isHighContrast?: boolean;
            foreground?: { value?: string };
            background?: { value?: string };
        };
        const isHC = !!palette?.isHighContrast;
        const hcFore = palette?.foreground?.value ?? "#000000";
        const hcBack = palette?.background?.value ?? "#ffffff";

        const barMode = this.enumValue(fs.columns.barMode, "grouped");
        const roundedCorners = !!fs.columns.roundedCorners.value;
        const cornerRadius = roundedCorners
            ? Math.max(0, Math.min(20, Number(fs.columns.cornerRadius.value) || 0)) : 0;
        const categoryGap = Math.max(0, Math.min(90, Number(fs.columns.categoryGap.value) || 30));
        const seriesGap = Math.max(0, Math.min(50, Number(fs.columns.seriesGap.value) || 10));
        // Single "Default color" override (empty = use the per-series palette).
        const barColorDefault = this.colorOf(fs.columns.barColor, "");
        const barAlpha = 1 - Math.max(0, Math.min(100, Number(fs.columns.barTransparency.value) || 0)) / 100;
        const borderShow = !!fs.columns.borderShow.value;
        const borderColorVal = this.colorOf(fs.columns.borderColor, "#000000");
        const borderWidthVal = Math.max(0, Math.min(10, Number(fs.columns.borderWidth.value) || 0));

        // ── Categories: leaf labels + optional parent group axis ─────────
        const categories = vm.categories.map(c => c.label);
        // Leaf axis shows only the INNERMOST level when a hierarchy is bound (the
        // joined label stays the unique category key); the parent row carries the
        // outer level. Keeps the two axis rows non-redundant + readable.
        const leafInner = new Map(vm.categories.map(c => [c.label, c.levels[c.levels.length - 1]]));
        this.totalCategories = categories.length;

        // Enter-animation gating: clear() makes every rebuild look "new" to
        // ECharts, so a 600ms intro would replay on every resize / format /
        // selection rebuild. Only animate when the dataset identity actually
        // changes (first paint, fixture/data swap, category-changing filter);
        // scroll never rebuilds, and same-data rebuilds stay instant.
        const dataSig = `${categories.length}|${vm.series.map(s => s.name).join(",")}|${categories.join(",")}`;
        const animateEnter = dataSig !== this.lastDataSig;
        this.lastDataSig = dataSig;

        // Scroll detection MUST precede the series build so the per-series
        // `animation` flag can be gated on it. ECharts resolves animation per
        // series via getShallow('animation'), so an explicit per-series `true`
        // (from animateEnter on first paint) OVERRIDES the root
        // `animation:false`. The wheel path snaps (resolved
        // animationDurationUpdate:0), but the dataZoom slider-DRAG roam
        // re-renders the bars with ECharts' 200ms `extraOpts` default tween,
        // bypassing that 0. Gating the series flag off while scrollable makes
        // isAnimationEnabled() === false → bars snap on every pan, every path.
        const showLegend = !!fs.legend.show.value
            && vm.series.length > 1
            && this.viewportW >= 220 && this.viewportH >= 160;
        const legendPos = this.enumValue(fs.legend.position, "top");
        const MIN_CAT_PX = 22;
        const reservedForLegendX = showLegend && (legendPos === "left" || legendPos === "right") ? 110 : 0;
        const plotW = Math.max(60, this.viewportW - 60 - reservedForLegendX);
        // Scrollbar can be disabled in the format pane — then ALL categories are
        // crammed into the viewport (thin bars) with no slider.
        const scrollbarOn = fs.scrollbar.show.value !== false;
        const needsScroll = scrollbarOn && categories.length * MIN_CAT_PX > plotW;
        this.scrollEnabled = needsScroll;
        // Scrollable + data changed → the intro animation will play; arm the
        // post-intro disable so zoom/pan stays snappy afterwards.
        this.pendingIntroDisable = animateEnter && needsScroll;
        let visibleCount = categories.length;
        if (needsScroll) {
            visibleCount = Math.max(2, Math.floor(plotW / MIN_CAT_PX));
        }

        // Outer-level labels: pick the OUTER hierarchy level (not the leaf).
        // For a 2-level hierarchy we use level[0] as the parent; the leaf is
        // levels[last] which already lives on the main axis.
        let parentAxisData: string[] | null = null;
        let parentDisplayName = "";
        if (vm.hierarchical && vm.categories[0].levels.length >= 2) {
            const outerIdx = 0;
            parentAxisData = vm.categories.map(c => c.levels[outerIdx]);
            parentDisplayName = vm.categoryFieldNames[outerIdx] ?? "";
            // Collapse repeated parent values down the axis — the second
            // axis renders each label, but visually merging duplicates makes
            // the hierarchy obvious. ECharts won't merge for us; the formatter
            // hides repeats by emitting "" for any label same as the previous.
            const seen: string[] = [];
            let last = "";
            for (const v of parentAxisData) { seen.push(v === last ? "" : v); last = v; }
            parentAxisData = seen;
            void parentDisplayName; // surfaced via tooltip context, not a name slot
        }

        // ── Per-series data (apply percent normalisation if needed) ──────
        const percentMode = barMode === "percent";
        const totals: number[] = percentMode
            ? categories.map((_, i) =>
                vm.series.reduce((sum, s) => sum + Math.max(0, s.points[i]?.value ?? 0), 0))
            : [];

        const hasSelection = this.localSelectedKeys.size > 0;
        const hasHighlights = vm.hasHighlights;

        const seriesOption = vm.series.map((s, sIdx) => {
            const data = s.points.map((p, i) => {
                const rawV = p.value;
                let value: number | null = rawV;
                if (percentMode) {
                    const total = totals[i];
                    value = (total > 0 && rawV != null) ? (Math.max(0, rawV) / total) * 100 : null;
                }

                // Dim un-selected / un-highlighted bars (combined with the
                // global bar transparency from the Columns card).
                let opacity = barAlpha;
                if (hasSelection) {
                    const k = idKey(p.selectionId);
                    opacity = (k && this.localSelectedKeys.has(k) ? 1 : 0.35) * barAlpha;
                } else if (hasHighlights) {
                    opacity = (p.highlight != null && p.highlight !== 0 ? 1 : 0.28) * barAlpha;
                }

                // Per-point override wins; else the single "Default color"; else the palette.
                const fill = p.colorOverride ?? (barColorDefault || s.color);
                // Round the VALUE end: top for positive (upward) bars, bottom for
                // negative (downward) bars — not always the top.
                const radius: [number, number, number, number] = value < 0
                    ? [0, 0, cornerRadius, cornerRadius]
                    : [cornerRadius, cornerRadius, 0, 0];
                return {
                    value,
                    itemStyle: {
                        color: isHC ? hcFore : fill,
                        opacity,
                        borderRadius: radius,
                        borderColor: isHC ? hcBack : (borderShow ? borderColorVal : "transparent"),
                        borderWidth: isHC ? 1 : (borderShow ? borderWidthVal : 0),
                    },
                };
            });

            const seriesEntry: Record<string, unknown> = {
                name: s.name,
                type: "bar",
                data,
                // Common to grouped + stacked:
                xAxisIndex: 0,
                yAxisIndex: 0,
                // Init "grow up" animation on a genuine data change
                // (animateEnter) — plays on scrollable charts too now that the
                // native dataZoom owns scroll. Updates snap instantly
                // (animationDurationUpdate:0) so zoom/pan never tweens.
                animation: animateEnter,
                animationDuration: 700,
                animationDurationUpdate: 0,
                animationDelay: (idx: number) => Math.min(idx * 14, 260),
                progressive: 0,
                progressiveThreshold: 600,
                largeThreshold: 200,
                barCategoryGap: `${categoryGap}%`,
            };
            if (barMode === "stacked" || barMode === "percent") {
                seriesEntry.stack = "total";
            } else {
                seriesEntry.barGap = `${seriesGap}%`;
            }
            // Data labels — show on the last (top) stacked series only when stacking
            const showLabels = !!fs.dataLabels.show.value;
            if (showLabels) {
                const stacked = barMode === "stacked" || barMode === "percent";
                const labelOnTop = (barMode === "grouped") || (sIdx === vm.series.length - 1);
                const dlPosMode = this.enumValue(fs.dataLabels.position, "auto");
                // auto → inside for stacked, top for grouped; explicit otherwise.
                const dlPosition = dlPosMode === "inside" ? "inside"
                    : dlPosMode === "outside" ? "top"
                        : (stacked ? "inside" : "top");
                const dlUnit = this.enumValue(fs.dataLabels.displayUnit, "auto");
                const dlDec = Math.max(0, Number(fs.dataLabels.decimalPlaces.value) || 0);
                seriesEntry.label = {
                    show: labelOnTop,
                    position: dlPosition,
                    fontFamily: (fs.dataLabels.fontFamily.value as string) || "Segoe UI",
                    fontSize: Math.max(7, Number(fs.dataLabels.fontSize.value) || 10),
                    fontWeight: fs.dataLabels.bold.value ? "bold" : "normal",
                    color: isHC ? hcFore : this.colorOf(fs.dataLabels.color, "#252423"),
                    formatter: (p: { value?: number | null }) =>
                        percentMode ? formatValue(p.value, "0%") : fmtWithUnit(p.value, dlUnit, dlDec),
                };
            }
            return seriesEntry;
        });

        // ── Total labels (stacked only) — a label-only scatter overlay at the
        //    top of each stack showing the category total. Percent mode totals
        //    are always 100%, so skip it there.
        if (barMode === "stacked" && fs.totalLabels.show.value) {
            const stackTotals = categories.map((_, i) =>
                vm.series.reduce((sum, s) => sum + (s.points[i]?.value ?? 0), 0));
            const tlUnit = this.enumValue(fs.totalLabels.displayUnit, "auto");
            const tlDec = Math.max(0, Number(fs.totalLabels.decimalPlaces.value) || 0);
            seriesOption.push({
                type: "scatter",
                xAxisIndex: 0,
                yAxisIndex: 0,
                data: stackTotals.map((t, i) => [i, t]),
                symbolSize: 0,
                silent: true,
                tooltip: { show: false },
                legendHoverLink: false,
                animation: false,
                label: {
                    show: true,
                    position: "top",
                    fontFamily: (fs.totalLabels.fontFamily.value as string) || "Segoe UI",
                    fontSize: Math.max(7, Number(fs.totalLabels.fontSize.value) || 10),
                    fontWeight: fs.totalLabels.bold.value ? "bold" : "normal",
                    color: isHC ? hcFore : this.colorOf(fs.totalLabels.color, "#252423"),
                    formatter: (p: { value?: number[] | number }) =>
                        fmtWithUnit(Array.isArray(p.value) ? p.value[1] : p.value, tlUnit, tlDec),
                },
                z: 5,
            } as Record<string, unknown>);
        }

        // ── Constant line (Analytics) — a horizontal reference line at a fixed
        //    value, drawn via markLine on the first bar series.
        if (fs.constantLine.show.value && seriesOption.length) {
            const clVal = Number(fs.constantLine.value.value) || 0;
            const clColor = isHC ? hcFore : this.colorOf(fs.constantLine.color, "#e44c4c");
            const clStyle = this.enumValue(fs.constantLine.lineStyle, "dashed");
            const clWidth = Math.max(0.5, Math.min(8, Number(fs.constantLine.lineWidth.value) || 2));
            const clShowLabel = !!fs.constantLine.showLabel.value;
            const clLabelText = (fs.constantLine.labelText.value as string) || "";
            (seriesOption[0] as Record<string, unknown>).markLine = {
                silent: true,
                symbol: "none",
                animation: false,
                lineStyle: { color: clColor, type: clStyle, width: clWidth },
                label: {
                    show: clShowLabel,
                    position: "insideEndTop",
                    color: isHC ? hcFore : this.colorOf(fs.constantLine.labelColor, "#e44c4c"),
                    formatter: () => clLabelText || abbreviate(clVal),
                },
                data: [{ yAxis: clVal }],
            };
        }

        // ── Legend ───────────────────────────────────────────────────────
        const legendColor = isHC ? hcFore : this.colorOf(fs.legend.color, "#605e5c");
        const legendFontSize = Math.max(8, Number(fs.legend.fontSize.value) || 11);
        const legendOpt: Record<string, unknown> = {
            show: showLegend,
            data: vm.series.map(s => s.name),
            textStyle: {
                color: legendColor,
                fontFamily: (fs.legend.fontFamily.value as string) || "Segoe UI",
                fontSize: legendFontSize,
                fontWeight: fs.legend.bold.value ? "bold" : "normal",
            },
            type: "scroll",
        };
        if (legendPos === "top") { legendOpt.top = 4; legendOpt.left = "center"; legendOpt.orient = "horizontal"; }
        else if (legendPos === "bottom") { legendOpt.bottom = 4; legendOpt.left = "center"; legendOpt.orient = "horizontal"; }
        else if (legendPos === "left") { legendOpt.left = 4; legendOpt.top = "middle"; legendOpt.orient = "vertical"; }
        else { legendOpt.right = 4; legendOpt.top = "middle"; legendOpt.orient = "vertical"; }

        // ── X axis (leaf labels) + optional parent group axis ────────────
        const xAxisColor = isHC ? hcFore : this.colorOf(fs.xAxis.color, "#605e5c");
        const xAxisFontSize = Math.max(8, Number(fs.xAxis.fontSize.value) || 11);
        const xRotate = this.enumValue(fs.xAxis.rotate, "auto");
        const rotateDeg = xRotate === "always" ? -45
            : xRotate === "never" ? 0
                : (categories.length > 8 ? -30 : 0);
        const xShow = !!fs.xAxis.show.value;

        // Vertical extents BELOW the leaf axis line — used to (a) offset the
        // parent group axis, (b) place the axis-name title clear of every
        // label row, and (c) reserve grid.bottom. `containLabel:true` reserves
        // the tick-label space automatically but does NOT account for the axis
        // name, so nameGap must be derived from these extents or the title
        // collides with the (rotated) labels.
        const xLabelPx = xShow
            ? (rotateDeg !== 0 ? Math.ceil(xAxisFontSize * 3.6) + 6 : xAxisFontSize + 12)
            : 4;
        // Parent group row must sit BELOW the (possibly rotated) leaf labels —
        // a fixed small offset lands the bold parent labels inside the rotated
        // leaf row. Clear the full leaf extent when rotated.
        const parentOffset = rotateDeg !== 0 ? xLabelPx + 4 : xAxisFontSize + 14;
        // Parent row height reserve follows the parent font when it's formatted
        // separately (a larger parent size must not clip or collide with the title).
        const parentFontPx = fs.xAxis.parentSeparate.value
            ? Math.max(8, Number(fs.xAxis.parentFontSize.value) || 12)
            : xAxisFontSize;
        const parentLabelPx = parentAxisData ? parentFontPx + 8 : 0;
        // Bottom-most extent of all label rows below the leaf axis line.
        const labelZoneBelowAxis = Math.max(xLabelPx, parentAxisData ? parentOffset + parentLabelPx : 0);

        // Anchor labels to absolute indices so they don't reshuffle as
        // dataZoom slides. With `interval: 0 + hideOverlap`, ECharts
        // re-evaluates which labels to draw against the *visible* window each
        // tick, so labels appear to jump. Setting `interval: N` selects by
        // `index % (N+1) === 0` against the absolute index — same categories
        // stay labeled across the entire scroll range.
        //
        // Compute density from the VISIBLE window, not the total dataset, so
        // the viewport is well-populated. Rotated labels need ~half the
        // horizontal slot of unrotated ones.
        const targetLabelPx = rotateDeg !== 0 ? 40 : 80;
        const visibleSlots = Math.max(2, Math.floor(plotW / targetLabelPx));
        const stride = needsScroll
            ? Math.max(1, Math.ceil(visibleCount / visibleSlots))
            : 1;
        const labelInterval = stride - 1;
        const leafAxisLabel = {
            show: xShow,
            color: xAxisColor,
            fontFamily: (fs.xAxis.fontFamily.value as string) || "Segoe UI",
            fontSize: xAxisFontSize,
            fontWeight: fs.xAxis.bold.value ? "bold" : "normal",
            interval: labelInterval,
            // Safety net: if the stride heuristic still leaves overlapping
            // labels (very long category names, small viewport), drop the
            // overlappers. Acts on the already-stable absolute-index set, so
            // the dropped labels stay dropped across scroll.
            hideOverlap: true,
            rotate: rotateDeg,
            width: 90,
            overflow: "truncate",
            // Hierarchical: show only the innermost level (parent row carries the
            // outer level). Flat: no formatter — the single label is shown as-is.
            ...(parentAxisData ? { formatter: (v: string) => leafInner.get(v) ?? v } : {}),
        };
        const xAxisCommon = {
            type: "category" as const,
            // Per-axis animation: false stops the label-position tween that
            // makes axis ticks "slide" between scroll ticks (compounds with
            // the top-level animationDurationUpdate:0 to make scroll feel
            // native).
            animation: false,
            // onZero:false pins the category axis line (+ labels + title) to
            // the BOTTOM of the plot for mixed +/- data — otherwise ECharts
            // draws it at the zero crossing (mid-plot) and the title geometry
            // becomes data-dependent. Bars still hang from zero regardless.
            axisLine: { show: xShow, onZero: false, lineStyle: { color: xAxisColor } },
            axisTick: { show: xShow, alignWithLabel: true },
            splitLine: { show: false },
        };
        // Vertical gridlines live on the leaf (main) x-axis splitLine.
        const vGridShow = !!fs.gridlines.verticalShow.value;
        const vGridColor = isHC ? hcFore : this.colorOf(fs.gridlines.verticalColor, "#e1dfdd");
        const vGridWidth = Math.max(0, Math.min(8, Number(fs.gridlines.verticalWidth.value) || 1));
        const vGridStyle = this.enumValue(fs.gridlines.verticalStyle, "solid");
        // Title defaults to the bound field name when shown without custom
        // text (native PBI behaviour) — the leaf category field for the X axis.
        const xTitleText = fs.xAxis.showTitle.value
            ? (String(fs.xAxis.title.value || "") || vm.categoryFieldNames[vm.categoryFieldNames.length - 1] || undefined)
            : undefined;
        const xTitleFontSize = Math.max(6, Number(fs.xAxis.titleFontSize.value) || 12);
        // Title sits below the bottom-most label row: the leaf labels, or the
        // parent group row (offset + its own label height) when present.
        const xTitleGap = labelZoneBelowAxis + Math.ceil(xTitleFontSize * 0.5) + 6;
        const xAxisList: unknown[] = [
            {
                ...xAxisCommon,
                data: categories,
                name: xTitleText,
                nameLocation: "middle",
                nameGap: xTitleGap,
                nameTextStyle: {
                    color: isHC ? hcFore : this.colorOf(fs.xAxis.titleColor, "#605e5c"),
                    fontSize: Math.max(6, Number(fs.xAxis.titleFontSize.value) || 12),
                },
                axisLabel: leafAxisLabel,
                splitLine: {
                    show: vGridShow,
                    lineStyle: { color: vGridColor, width: vGridWidth, type: vGridStyle },
                },
            },
        ];
        if (parentAxisData) {
            // Parent (outer) row: when "Format parent level separately" is on it
            // gets its own font/size/bold/colour; otherwise it inherits the leaf
            // format and is force-bold (the prior behaviour).
            const parentAxisLabel = fs.xAxis.parentSeparate.value
                ? {
                    ...leafAxisLabel,
                    fontFamily: (fs.xAxis.parentFontFamily.value as string) || "Segoe UI",
                    fontSize: parentFontPx,
                    fontWeight: fs.xAxis.parentBold.value ? "bold" : "normal",
                    color: isHC ? hcFore : this.colorOf(fs.xAxis.parentColor, "#605e5c"),
                    width: undefined,
                    overflow: "none",
                }
                : {
                    ...leafAxisLabel,
                    fontWeight: "bold",
                    width: undefined,
                    overflow: "none",
                };
            xAxisList.push({
                ...xAxisCommon,
                data: parentAxisData,
                position: "bottom",
                offset: parentOffset,
                axisLabel: parentAxisLabel,
                // Shadow band only on the leaf axis — avoid double rendering.
                axisPointer: { show: false },
            });
        }

        // ── Y axis ──────────────────────────────────────────────────────
        const yAxisColor = isHC ? hcFore : this.colorOf(fs.yAxis.color, "#605e5c");
        const yAxisFontSize = Math.max(8, Number(fs.yAxis.fontSize.value) || 11);
        const gridColor = isHC ? hcFore : this.colorOf(fs.gridlines.color, "#e1dfdd");
        const gridWidth = Math.max(0, Math.min(8, Number(fs.gridlines.width.value) || 1));
        const gridStyle = this.enumValue(fs.gridlines.style, "solid");
        const yShow = !!fs.yAxis.show.value;
        // Range: blank / "Auto" / non-numeric → auto (undefined). Percent mode pins 0–100.
        const parseRange = (slot: { value?: unknown }): number | undefined => {
            const s = String(slot?.value ?? "").trim();
            if (!s || s.toLowerCase() === "auto") return undefined;
            const n = Number(s);
            return Number.isFinite(n) ? n : undefined;
        };
        const yUnit = this.enumValue(fs.yAxis.displayUnit, "auto");
        const yDec = Math.max(0, Number(fs.yAxis.decimalPlaces.value) || 0);
        // Y-axis title defaults to the measure name when shown without text.
        const yTitleText = fs.yAxis.showTitle.value
            ? (String(fs.yAxis.title.value || "") || vm.valueDisplayName || undefined)
            : undefined;
        const yTitleFontSize = Math.max(6, Number(fs.yAxis.titleFontSize.value) || 12);
        // Width reserved for the y tick labels ("250k", "-60k") and (when
        // shown) the rotated title — drives the exact gridLeft below.
        const yLabelBand = yShow ? Math.max(34, Math.ceil(yAxisFontSize * 3.6)) : 4;
        const yNameGap = yLabelBand + Math.ceil(yTitleFontSize * 0.5) + 6;
        const yAxisOpt = {
            type: "value" as const,
            show: yShow,
            name: yTitleText,
            nameLocation: "middle" as const,
            nameGap: yNameGap,
            nameTextStyle: {
                color: isHC ? hcFore : this.colorOf(fs.yAxis.titleColor, "#605e5c"),
                fontSize: yTitleFontSize,
            },
            axisLabel: {
                show: yShow,
                color: yAxisColor,
                fontFamily: (fs.yAxis.fontFamily.value as string) || "Segoe UI",
                fontSize: yAxisFontSize,
                fontWeight: fs.yAxis.bold.value ? "bold" : "normal",
                hideOverlap: true,
                formatter: (v: number) => percentMode ? `${v.toFixed(0)}%` : fmtWithUnit(v, yUnit, yDec),
            },
            axisLine: { show: yShow, lineStyle: { color: yAxisColor } },
            splitLine: {
                show: !!fs.gridlines.show.value,
                lineStyle: { color: gridColor, width: gridWidth, type: gridStyle },
            },
            max: percentMode ? 100 : parseRange(fs.yAxis.rangeEnd),
            min: percentMode ? 0 : parseRange(fs.yAxis.rangeStart),
        };

        // ── Plot-rect margins (computed before dataZoom so the slider can
        //    align with the plot) — EXACT manual reserves, containLabel:false.
        // containLabel:true reserves only the tick LABELS (not the axis name),
        // and its hidden reserve varies with label LENGTH — so a fixed nameGap
        // either double-reserves (big empty gap above the slider) or
        // under-reserves (title clipped, e.g. short labels on negative-value
        // charts). Reserving every band ourselves makes the plot fill the cell
        // and the title land in the same place regardless of data. (Same
        // approach as waterfall-chart.)
        const legendTop = showLegend && legendPos === "top" ? legendFontSize + 14 : 0;
        const legendBot = showLegend && legendPos === "bottom" ? legendFontSize + 14 : 0;
        const legendLeft = showLegend && legendPos === "left" ? reservedForLegendX : 0;
        const legendRight = showLegend && legendPos === "right" ? reservedForLegendX : 0;
        const sliderPx = needsScroll ? 32 : 0;
        const xTitleBand = xTitleText ? xTitleFontSize + 8 : 0;
        const outerPad = 8;
        const gridLeft = outerPad + legendLeft + yLabelBand + (yTitleText ? yTitleFontSize + 6 : 0);
        const gridRight = outerPad + legendRight;
        const gridTop = outerPad + legendTop;
        const gridBottom = outerPad + legendBot + sliderPx + labelZoneBelowAxis + xTitleBand;
        const sliderLeft = gridLeft;
        const sliderRight = gridRight;

        // ── dataZoom: native zoom+pan slider + inside (mix-zoom-on-value) ───
        const dataZoom: unknown[] = [];
        if (needsScroll) {
            // Initial window shows ~visibleCount of the total; the live
            // start/end % is carried across rebuilds (captureScroll) so the
            // user's zoom + pan survive a peer cross-filter / resize / format
            // change instead of snapping back.
            const defEnd = Math.min(100, (visibleCount / Math.max(1, categories.length)) * 100);
            const startPct = this.scrollStartPct >= 0 ? this.scrollStartPct : 0;
            const endPct = this.scrollEndPct >= 0 ? this.scrollEndPct : defEnd;
            this.scrollStartPct = startPct;
            this.scrollEndPct = endPct;
            const xAxisIdx = parentAxisData ? [0, 1] : [0];
            dataZoom.push({
                ...SLIDER_STYLE,
                id: "dz",
                xAxisIndex: xAxisIdx,
                orient: "horizontal",
                bottom: 4,
                height: 24,
                left: sliderLeft,
                right: sliderRight,
                start: startPct,
                end: endPct,
            });
            // Inside zoom: wheel zooms (centred on cursor), drag pans — the
            // same gesture set as the reference example.
            dataZoom.push({
                type: "inside",
                id: "dz-inside",
                xAxisIndex: xAxisIdx,
                filterMode: "filter",
                realtime: true,
                start: startPct,
                end: endPct,
                zoomOnMouseWheel: true,
                moveOnMouseMove: true,
                moveOnMouseWheel: false,
                throttle: 0,
            });
        } else {
            this.scrollStartPct = -1;
            this.scrollEndPct = -1;
        }

        // ── Grid (plot rect) — margins computed above ────────────────────
        const plotAreaShow = !!fs.plotArea.show.value;
        const plotAreaAlpha = 1 - Math.max(0, Math.min(100, Number(fs.plotArea.transparency.value) || 0)) / 100;
        const grid = {
            left: gridLeft,
            right: gridRight,
            top: gridTop,
            bottom: gridBottom,
            containLabel: false,
            show: plotAreaShow,
            backgroundColor: plotAreaShow
                ? hexToRgba(this.colorOf(fs.plotArea.color, "#ffffff"), plotAreaAlpha) : "transparent",
            borderWidth: 0,
        };

        // ── Tooltip — axis-trigger for intent hovering ──────────────────
        //
        // Hovering anywhere over a category column reveals the tooltip
        // for that column (and the shadow band paints over it), instead of
        // requiring a pixel-perfect hit on a bar. ECharts auto-syncs the
        // shadow band to the cursor's category, so we drop the per-axis
        // axisPointer config in favour of the tooltip-level one below.
        // Tooltip style: "custom" shows the ECharts HTML box; "powerbi" hides it
        // (showContent:false) so the native host tooltip is used instead — the
        // axisPointer shadow band stays either way. Avoids the double tooltip.
        const tooltipStyle = this.enumValue(fs.tooltip.style, "custom");
        const tooltipOpt = {
            trigger: "axis" as const,
            showContent: tooltipStyle !== "powerbi",
            axisPointer: {
                type: "shadow" as const,
                // Derive the hover band from the HC foreground (translucent) so
                // it respects the user's high-contrast theme rather than a
                // hard-coded black.
                shadowStyle: isHC ? { color: hcFore, opacity: 0.18 } : { color: "rgba(0,0,0,0.10)" },
                label: { show: false },
            },
            confine: true,
            transitionDuration: 0,
            backgroundColor: isHC ? hcBack : "rgba(33,37,41,0.96)",
            borderColor: isHC ? hcFore : "transparent",
            borderWidth: isHC ? 1 : 0,
            extraCssText: "transition: opacity .16s ease;",
            textStyle: { color: isHC ? hcFore : "#ffffff", fontSize: 12 },
            formatter: (p: unknown) => this.tooltipHtml(p, vm, percentMode),
        };

        return {
            // Init "grow up" animation on a genuine data change (animateEnter),
            // including scrollable charts — the NATIVE slider+inside dataZoom
            // owns scroll and re-renders as UPDATEs, which snap via
            // animationDurationUpdate:0 (no tween). The old custom-wheel slider
            // needed animation off entirely; the native dataZoom doesn't.
            animation: animateEnter,
            animationDuration: 700,
            animationDurationUpdate: 0,
            animationEasing: "cubicOut" as const,
            animationEasingUpdate: "cubicInOut" as const,
            stateAnimation: { duration: 250, easing: "cubicOut" as const },
            grid,
            legend: legendOpt,
            tooltip: tooltipOpt,
            xAxis: xAxisList,
            yAxis: yAxisOpt,
            series: seriesOption,
            ...(dataZoom.length ? { dataZoom } : {}),
        } as unknown as EChartsOption;
    }

    // ─── Tooltip ─────────────────────────────────────────────────────────

    private tooltipHtml(raw: unknown, vm: ColumnViewModel, percentMode: boolean): string {
        // Axis trigger passes an ARRAY of items (one per series at that x).
        // Item trigger passes a single object. Normalise to array.
        const arr = Array.isArray(raw)
            ? (raw as Array<{ seriesIndex?: number; dataIndex?: number; color?: string; value?: unknown }>)
            : [raw as { seriesIndex?: number; dataIndex?: number; color?: string; value?: unknown }];
        if (arr.length === 0) return "";

        const dIdx = arr[0].dataIndex ?? -1;
        const cat = vm.categories[dIdx];
        if (!cat) return "";

        let header = `<div style="font-weight:600;margin-bottom:4px">${escapeHtml(cat.label)}</div>`;
        if (vm.hierarchical && cat.levels.length > 1) {
            header = `<div style="font-weight:600;margin-bottom:4px">`
                + cat.levels.map(l => escapeHtml(l)).join(" › ")
                + `</div>`;
        }

        // Per-series rows — skip series with no value at this index.
        const rows: string[] = [];
        let firstPoint: ColumnSeries["points"][number] | undefined;
        for (const item of arr) {
            const sIdx = item.seriesIndex ?? -1;
            const series = vm.series[sIdx];
            if (!series) continue;
            const point = series.points[dIdx];
            if (!point) continue;
            if (!firstPoint) firstPoint = point;
            const rawVal = point.value;
            if (rawVal == null) continue;
            const valStr = formatValue(rawVal, vm.valueFormat);
            const swatch = `<span style="display:inline-block;width:8px;height:8px;border-radius:2px;`
                + `background:${escapeHtml(item.color || series.color)};margin-right:6px;vertical-align:middle"></span>`;
            let line = `<div style="line-height:1.5">${swatch}`
                + `<span style="opacity:.85">${escapeHtml(series.name)}: </span>`
                + `<b>${escapeHtml(valStr)}</b>`;
            if (percentMode && Number.isFinite(rawVal)) {
                const total = vm.series.reduce((sum, s) => sum + Math.max(0, s.points[dIdx]?.value ?? 0), 0);
                const pct = total > 0 ? (Math.max(0, rawVal) / total) * 100 : 0;
                line += ` <span style="opacity:.6">(${pct.toFixed(1)}%)</span>`;
            }
            line += `</div>`;
            rows.push(line);
        }

        let extra = "";
        if (firstPoint?.tooltips?.length) {
            for (const t of firstPoint.tooltips) {
                const tv = t.value;
                const tStr = tv == null ? "—"
                    : typeof tv === "number" ? formatValue(tv, t.format) : String(tv);
                extra += `<div style="opacity:.7;margin-top:2px">${escapeHtml(t.displayName)}: `
                    + `<b>${escapeHtml(tStr)}</b></div>`;
            }
        }
        return header + rows.join("") + extra;
    }

    // ─── Host wiring ─────────────────────────────────────────────────────

    private bindChartEvents(): void {
        const chart = this.chart;
        if (!chart) return;

        chart.on("click", { seriesType: "bar" }, (params: unknown) => {
            if (!this.allowInteractions || !this.viewModel) return;
            const p = params as { seriesIndex?: number; dataIndex?: number; event?: { event?: MouseEvent } };
            const s = this.viewModel.series[p.seriesIndex ?? -1];
            const point = s?.points[p.dataIndex ?? -1];
            const id = point?.selectionId;
            if (!id) return;
            const evt = p.event?.event;
            const multi = !!(evt?.ctrlKey || evt?.metaKey);
            this.handleSelect(id, multi);
        });

        chart.on("contextmenu", { seriesType: "bar" }, (params: unknown) => {
            const p = params as { seriesIndex?: number; dataIndex?: number; event?: { event?: MouseEvent } };
            const s = this.viewModel?.series[p.seriesIndex ?? -1];
            const point = s?.points[p.dataIndex ?? -1];
            const id = point?.selectionId ?? null;
            const evt = p.event?.event;
            try { evt?.preventDefault(); } catch { /* ignore */ }
            this.selectionManager.showContextMenu(id ?? {}, { x: evt?.clientX ?? 0, y: evt?.clientY ?? 0 });
        });

        // Cross-filter on legend item click — ECharts' legend toggles series
        // visibility by default; revert that and route to selectionManager.
        chart.on("legendselectchanged", (raw: unknown) => {
            // Reverting the toggle below re-fires this event — guard against it.
            if (this.legendSyncing) return;
            const p = raw as { name?: string; selected?: Record<string, boolean> };
            // Undo ECharts' show/hide so the series never actually hides.
            this.legendSyncing = true;
            for (const k of Object.keys(p.selected ?? {})) {
                if (!p.selected![k]) chart.dispatchAction({ type: "legendSelect", name: k });
            }
            this.legendSyncing = false;
            if (!this.allowInteractions || !this.viewModel || !p.name) return;
            const series = this.viewModel.series.find(s => s.name === p.name);
            const id = series?.legendSelectionId;
            if (!id) return;
            // legendselectchanged carries no MouseEvent; use the ctrl/⌘ state
            // captured on the preceding mousedown so legend clicks can multi-select.
            this.handleSelect(id, this.lastPointerMulti);
        });

        const zr = (chart as unknown as { getZr?: () => {
            on: (ev: string, h: (e: { target?: unknown; event?: MouseEvent }) => void) => void;
        } }).getZr?.();
        if (zr) {
            // Capture the multi-select modifier on mousedown — both the click
            // and the legendselectchanged handlers read it (the latter has no
            // MouseEvent of its own).
            zr.on("mousedown", (e) => {
                const me = e.event as MouseEvent | undefined;
                this.lastPointerMulti = !!(me?.ctrlKey || me?.metaKey);
            });
            zr.on("click", (e) => {
                if (e.target) return;
                this.localSelectedKeys = new Set();
                this.selectionManager.clear().then(() => this.applySelectionDim());
                this.applySelectionDim();
            });
            zr.on("contextmenu", (e) => {
                if (e.target) return;
                try { e.event?.preventDefault(); } catch { /* ignore */ }
                this.selectionManager.showContextMenu({}, { x: e.event?.clientX ?? 0, y: e.event?.clientY ?? 0 });
            });
            // Wheel zoom is owned by the native `inside` dataZoom — no manual handler.
        }
        this.chartEl.addEventListener("contextmenu", (e) => e.preventDefault());

        // Mirror ITooltipService so the host's tooltip-context (drill-through,
        // report-page tooltips) stays wired even though the visible UI is the
        // ECharts native tooltip.
        if (this.tooltipService) {
            chart.on("mouseover", { seriesType: "bar" }, (raw: unknown) => {
                // Only drive the native host tooltip in "powerbi" mode — otherwise
                // it double-renders with the ECharts tooltip (read live; bound once).
                if (this.enumValue(this.formattingSettings?.tooltip?.style, "custom") !== "powerbi") return;
                if (!this.viewModel || !this.tooltipService) return;
                const p = raw as { seriesIndex?: number; dataIndex?: number; event?: { event?: MouseEvent } };
                const s = this.viewModel.series[p.seriesIndex ?? -1];
                const point = s?.points[p.dataIndex ?? -1];
                if (!s || !point) return;
                const rect = this.chartEl.getBoundingClientRect();
                const evt = p.event?.event;
                this.tooltipService.show({
                    coordinates: [
                        (evt?.clientX ?? rect.left) - rect.left,
                        (evt?.clientY ?? rect.top) - rect.top,
                    ],
                    isTouchEvent: false,
                    dataItems: [
                        { displayName: this.viewModel.categories[p.dataIndex ?? -1]?.label ?? "", value: "" },
                        { displayName: s.name, value: formatValue(point.value, this.viewModel.valueFormat), color: s.color },
                    ],
                    identities: point.selectionId ? [point.selectionId] : [],
                });
            });
            chart.on("mouseout", () => {
                this.tooltipService?.hide({ immediately: false, isTouchEvent: false });
            });
        }
    }

    private handleSelect(id: ISelectionId, multi: boolean): void {
        if (!this.allowInteractions) return;
        // Monotonic guard: rapid clicks resolve out of order; an older .then
        // landing last would clobber the newer selection and flicker the dim.
        const seq = ++this.clickSeq;
        const k = idKey(id);
        const already = this.localSelectedKeys.size === 1 && this.localSelectedKeys.has(k);
        if (already && !multi) {
            // Toggling the last selection off must route through clear() —
            // select() can never produce an empty selection.
            this.localSelectedKeys = new Set();
            this.applySelectionDim();
            this.selectionManager.clear().then(() => { if (seq === this.clickSeq) this.applySelectionDim(); });
            return;
        }
        if (multi) this.localSelectedKeys.add(k); else this.localSelectedKeys = new Set([k]);
        this.applySelectionDim();
        this.selectionManager.select(id, multi).then((ids) => {
            if (seq !== this.clickSeq) return;
            this.localSelectedKeys = new Set((ids as ISelectionId[]).map(idKey));
            this.applySelectionDim();
        });
    }

    /** Read the live dataZoom window (start/end %) before a rebuild, so the
     *  user's slider zoom + pan and inside wheel-zoom survive update(). The
     *  native slider + inside dataZoom own the interaction — there is no manual
     *  wheel handler. */
    private captureScroll(): void {
        const chart = this.chart;
        if (!chart || !this.scrollEnabled) return;
        const dz = (chart as unknown as {
            getModel?: () => { queryComponents?: (q: object) => { option?: unknown }[] };
        }).getModel?.()?.queryComponents?.({ mainType: "dataZoom", id: "dz" })?.[0];
        const o = dz?.option as { start?: number; end?: number } | undefined;
        if (o && typeof o.start === "number" && typeof o.end === "number") {
            this.scrollStartPct = o.start;
            this.scrollEndPct = o.end;
        }
    }

    // ─── Keyboard navigation ─────────────────────────────────────────────

    private handleKeyDown(e: KeyboardEvent): void {
        const vm = this.viewModel;
        const chart = this.chart;
        if (!vm || vm.isEmpty || !chart) return;
        const nCat = vm.categories.length;
        const nSer = vm.series.length;
        if (nCat === 0) return;

        switch (e.key) {
            case "ArrowRight":
                this.focusedCat = this.focusedCat < 0 ? 0 : Math.min(nCat - 1, this.focusedCat + 1);
                break;
            case "ArrowLeft":
                this.focusedCat = this.focusedCat <= 0 ? 0 : this.focusedCat - 1;
                break;
            case "ArrowUp":
                if (this.focusedCat < 0) this.focusedCat = 0;
                this.focusedSeries = Math.max(0, this.focusedSeries - 1);
                break;
            case "ArrowDown":
                if (this.focusedCat < 0) this.focusedCat = 0;
                this.focusedSeries = Math.min(nSer - 1, this.focusedSeries + 1);
                break;
            case "Enter":
            case " ": {
                if (this.focusedCat < 0) this.focusedCat = 0;
                const pt = vm.series[this.focusedSeries]?.points[this.focusedCat];
                if (pt?.selectionId) this.handleSelect(pt.selectionId, e.shiftKey || e.ctrlKey || e.metaKey);
                e.preventDefault();
                return;
            }
            case "Escape":
                this.localSelectedKeys = new Set();
                this.applySelectionDim();
                this.selectionManager.clear().then(() => this.applySelectionDim());
                e.preventDefault();
                return;
            default:
                return;   // don't swallow other keys
        }
        e.preventDefault();
        this.ensureCatVisible(this.focusedCat);
        this.highlightFocused();
    }

    /** Pan the dataZoom window (keeping its current zoom width) so the focused
     *  category is on screen — works in the start/end % space. */
    private ensureCatVisible(cat: number): void {
        const chart = this.chart;
        if (!chart || !this.scrollEnabled || cat < 0 || this.totalCategories <= 1) return;
        const dz = (chart as unknown as {
            getModel?: () => { queryComponents?: (q: object) => { option?: unknown }[] };
        }).getModel?.()?.queryComponents?.({ mainType: "dataZoom", id: "dz" })?.[0];
        const o = dz?.option as { start?: number; end?: number } | undefined;
        if (!o || typeof o.start !== "number" || typeof o.end !== "number") return;
        const width = o.end - o.start;
        const catPct = (cat / (this.totalCategories - 1)) * 100;
        if (catPct >= o.start && catPct <= o.end) return;
        let start = catPct < o.start ? catPct : catPct - width;
        start = Math.max(0, Math.min(100 - width, start));
        this.scrollStartPct = start;
        this.scrollEndPct = start + width;
        chart.dispatchAction({ type: "dataZoom", dataZoomId: "dz", start, end: start + width });
    }

    /** Emphasis ring on the focused bar (keyboard focus indicator). */
    private highlightFocused(): void {
        const chart = this.chart;
        if (!chart || this.focusedCat < 0) return;
        try {
            chart.dispatchAction({ type: "downplay" });
            chart.dispatchAction({ type: "highlight", seriesIndex: this.focusedSeries, dataIndex: this.focusedCat });
        } catch { /* ignore */ }
    }

    /** Re-emit per-bar opacity from `localSelectedKeys` / highlights. */
    private applySelectionDim(): void {
        const chart = this.chart;
        const vm = this.viewModel;
        if (!chart || !vm) return;
        const hasSelection = this.localSelectedKeys.size > 0;
        const hasHighlights = vm.hasHighlights;
        // ECharts replaces the whole series.data array (per-item merge doesn't
        // happen), so we must re-emit every field — not just opacity.
        const fs = this.formattingSettings;
        const palette = this.host.colorPalette as { isHighContrast?: boolean; foreground?: { value?: string }; background?: { value?: string } };
        const isHC = !!palette?.isHighContrast;
        const hcFore = palette?.foreground?.value ?? "#000000";
        const hcBack = palette?.background?.value ?? "#ffffff";
        const roundedCorners = !!fs.columns.roundedCorners.value;
        const cornerRadius = roundedCorners ? Math.max(0, Math.min(20, Number(fs.columns.cornerRadius.value) || 0)) : 0;
        const barColorDefault = this.colorOf(fs.columns.barColor, "");
        const barAlpha = 1 - Math.max(0, Math.min(100, Number(fs.columns.barTransparency.value) || 0)) / 100;
        const borderShow = !!fs.columns.borderShow.value;
        const borderColorVal = this.colorOf(fs.columns.borderColor, "#000000");
        const borderWidthVal = Math.max(0, Math.min(10, Number(fs.columns.borderWidth.value) || 0));
        const barMode = this.enumValue(fs.columns.barMode, "grouped");
        const percentMode = barMode === "percent";
        const totals: number[] = percentMode
            ? vm.categories.map((_, i) =>
                vm.series.reduce((sum, s) => sum + Math.max(0, s.points[i]?.value ?? 0), 0))
            : [];
        try {
            // Re-emit only the bar series (indices 0..n-1); the optional total-label
            // scatter + markLine live at higher indices and merge-survive untouched.
            chart.setOption({
                series: vm.series.map((s) => ({
                    data: s.points.map((p, i) => {
                        const rawV = p.value;
                        let value: number | null = rawV;
                        if (percentMode) {
                            const total = totals[i];
                            value = (total > 0 && rawV != null) ? (Math.max(0, rawV) / total) * 100 : null;
                        }
                        let opacity = barAlpha;
                        if (hasSelection) {
                            const k = idKey(p.selectionId);
                            opacity = (k && this.localSelectedKeys.has(k) ? 1 : 0.35) * barAlpha;
                        } else if (hasHighlights) {
                            opacity = (p.highlight != null && p.highlight !== 0 ? 1 : 0.28) * barAlpha;
                        }
                        const fill = p.colorOverride ?? (barColorDefault || s.color);
                        return {
                            value,
                            itemStyle: {
                                color: isHC ? hcFore : fill,
                                opacity,
                                borderRadius: (typeof value === "number" && value < 0)
                                    ? [0, 0, cornerRadius, cornerRadius]
                                    : [cornerRadius, cornerRadius, 0, 0],
                                borderColor: isHC ? hcBack : (borderShow ? borderColorVal : "transparent"),
                                borderWidth: isHC ? 1 : (borderShow ? borderWidthVal : 0),
                            },
                        } as unknown as number;
                    }),
                })),
            }, false);
        } catch { /* chart may be torn down */ }
    }
}

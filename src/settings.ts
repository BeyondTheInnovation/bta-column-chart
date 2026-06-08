"use strict";

/**
 * FormattingSettingsModel for the column-chart visual (ECharts engine).
 *
 * Cards mirror `capabilities.json` one-for-one. Card `name` ↔ `objects` key and
 * slice `name` ↔ property key must match `capabilities.json` exactly or the
 * setting silently never round-trips.
 *
 * NOTE (in-place upgrade from the D3 column-chart): the card/slice `name`
 * strings are kept BACKWARD-COMPATIBLE with the shipped D3 visual's object
 * names (`categoryColors`, `barSettings`, `xAxisSettings`, `yAxisSettings`,
 * `legend`, `gridlines`, `dataLabels`, `totalLabels`, `plotArea`,
 * `constantLine`) so existing reports keep their saved formatting after the
 * engine swap. The JS property names stay engine-friendly so visual.ts reads
 * naturally.
 */

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import Card = formattingSettings.SimpleCard;
import Model = formattingSettings.Model;
import Slice = formattingSettings.Slice;

const TEXT_FONT = "\"Segoe UI\", wf_segoe-ui_normal, helvetica, arial, sans-serif";

// `instanceKind` value that enables the `fx` (conditional / rule-based) button
// on a ColorPicker slice.
const CONSTANT_OR_RULE =
    (powerbi as unknown as { VisualEnumerationInstanceKinds?: { ConstantOrRule?: number } })
        ?.VisualEnumerationInstanceKinds?.ConstantOrRule ?? 3;

const DISPLAY_UNIT_ITEMS = [
    { value: "auto", displayName: "Auto" },
    { value: "none", displayName: "None" },
    { value: "thousands", displayName: "Thousands (K)" },
    { value: "millions", displayName: "Millions (M)" },
    { value: "billions", displayName: "Billions (B)" },
];

const LINE_STYLE_ITEMS = [
    { value: "solid", displayName: "Solid" },
    { value: "dashed", displayName: "Dashed" },
    { value: "dotted", displayName: "Dotted" },
];

/** Dynamic "Data colors" card — slices are rebuilt every render in
 *  `getFormattingModel`, one ColorPicker per distinct series / category.
 *  Object name `categoryColors` matches the shipped D3 visual. */
class DataColorsCard extends Card {
    name = "categoryColors";
    displayName = "Data colors";
    slices: Slice[] = [];
}

class ColumnsCard extends Card {
    name = "barSettings";
    displayName = "Columns";

    barMode = new formattingSettings.ItemDropdown({
        name: "barMode",
        displayName: "Bar mode",
        items: [
            { value: "grouped", displayName: "Grouped" },
            { value: "stacked", displayName: "Stacked" },
            { value: "percent", displayName: "100% Stacked" },
        ],
        value: { value: "grouped", displayName: "Grouped" },
    });
    roundedCorners = new formattingSettings.ToggleSwitch({
        name: "roundedCorners", displayName: "Rounded corners", value: true,
    });
    cornerRadius = new formattingSettings.NumUpDown({
        name: "cornerRadius", displayName: "Corner radius (px)", value: 2,
    });
    categoryGap = new formattingSettings.NumUpDown({
        name: "spaceBetweenCategories", displayName: "Space between categories (%)", value: 30,
    });
    seriesGap = new formattingSettings.NumUpDown({
        name: "spaceBetweenSeries", displayName: "Space between series (%)", value: 10,
    });
    barColor = new formattingSettings.ColorPicker({
        name: "barColor", displayName: "Default color",
        value: { value: "" },
        instanceKind: CONSTANT_OR_RULE,
    });
    barTransparency = new formattingSettings.NumUpDown({
        name: "barTransparency", displayName: "Transparency (%)", value: 0,
    });
    borderShow = new formattingSettings.ToggleSwitch({
        name: "borderShow", displayName: "Border", value: false,
    });
    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor", displayName: "Border color", value: { value: "#000000" },
    });
    borderWidth = new formattingSettings.NumUpDown({
        name: "borderWidth", displayName: "Border width (px)", value: 1,
    });

    slices: Slice[] = [
        this.barMode, this.roundedCorners, this.cornerRadius, this.categoryGap, this.seriesGap,
        this.barColor, this.barTransparency, this.borderShow, this.borderColor, this.borderWidth,
    ];
}

class XAxisCard extends Card {
    name = "xAxisSettings";
    displayName = "X-axis";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", value: true });
    fontFamily = new formattingSettings.FontPicker({ name: "fontFamily", displayName: "Font", value: TEXT_FONT });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Size", value: 11 });
    bold = new formattingSettings.ToggleSwitch({ name: "bold", displayName: "Bold", value: false });
    color = new formattingSettings.ColorPicker({
        name: "color", displayName: "Color", value: { value: "#605e5c" }, instanceKind: CONSTANT_OR_RULE,
    });
    rotate = new formattingSettings.ItemDropdown({
        name: "rotateLabels",
        displayName: "Rotate labels",
        items: [
            { value: "auto", displayName: "Auto" },
            { value: "always", displayName: "Always (-45°)" },
            { value: "never", displayName: "Never (0°)" },
        ],
        value: { value: "auto", displayName: "Auto" },
    });
    // Per-level formatting for a hierarchical (two-level) X-axis. When off, the
    // parent (outer) label row inherits the leaf format and is rendered bold —
    // today's behaviour. When on, the parent row gets its own font/size/bold/
    // colour so the two levels can be told apart at a glance.
    parentSeparate = new formattingSettings.ToggleSwitch({ name: "parentSeparate", displayName: "Format parent level separately", value: false });
    parentFontFamily = new formattingSettings.FontPicker({ name: "parentFontFamily", displayName: "Parent font", value: TEXT_FONT });
    parentFontSize = new formattingSettings.NumUpDown({ name: "parentFontSize", displayName: "Parent size", value: 12 });
    parentBold = new formattingSettings.ToggleSwitch({ name: "parentBold", displayName: "Parent bold", value: true });
    parentColor = new formattingSettings.ColorPicker({
        name: "parentColor", displayName: "Parent color", value: { value: "#605e5c" }, instanceKind: CONSTANT_OR_RULE,
    });

    showTitle = new formattingSettings.ToggleSwitch({ name: "showTitle", displayName: "Title", value: false });
    title = new formattingSettings.TextInput({ name: "titleText", displayName: "Title text", value: "", placeholder: "" });
    titleFontSize = new formattingSettings.NumUpDown({ name: "titleFontSize", displayName: "Title size", value: 12 });
    titleColor = new formattingSettings.ColorPicker({ name: "titleColor", displayName: "Title color", value: { value: "#605e5c" } });

    topLevelSlice = this.show;
    slices: Slice[] = [
        this.fontFamily, this.fontSize, this.bold, this.color, this.rotate,
        this.parentSeparate, this.parentFontFamily, this.parentFontSize, this.parentBold, this.parentColor,
        this.showTitle, this.title, this.titleFontSize, this.titleColor,
    ];
}

class YAxisCard extends Card {
    name = "yAxisSettings";
    displayName = "Y-axis";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", value: true });
    fontFamily = new formattingSettings.FontPicker({ name: "fontFamily", displayName: "Font", value: TEXT_FONT });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Size", value: 11 });
    bold = new formattingSettings.ToggleSwitch({ name: "bold", displayName: "Bold", value: false });
    color = new formattingSettings.ColorPicker({
        name: "color", displayName: "Color", value: { value: "#605e5c" }, instanceKind: CONSTANT_OR_RULE,
    });
    // Range — blank string = Auto (matches the D3 visual's Auto sentinel).
    rangeStart = new formattingSettings.TextInput({ name: "rangeStart", displayName: "Start", value: "", placeholder: "Auto" });
    rangeEnd = new formattingSettings.TextInput({ name: "rangeEnd", displayName: "End", value: "", placeholder: "Auto" });
    displayUnit = new formattingSettings.ItemDropdown({
        name: "displayUnit", displayName: "Display units",
        items: DISPLAY_UNIT_ITEMS, value: { value: "auto", displayName: "Auto" },
    });
    decimalPlaces = new formattingSettings.NumUpDown({ name: "decimalPlaces", displayName: "Decimal places", value: 0 });
    showTitle = new formattingSettings.ToggleSwitch({ name: "showTitle", displayName: "Title", value: false });
    title = new formattingSettings.TextInput({ name: "titleText", displayName: "Title text", value: "", placeholder: "" });
    titleFontSize = new formattingSettings.NumUpDown({ name: "titleFontSize", displayName: "Title size", value: 12 });
    titleColor = new formattingSettings.ColorPicker({ name: "titleColor", displayName: "Title color", value: { value: "#605e5c" } });

    topLevelSlice = this.show;
    slices: Slice[] = [
        this.fontFamily, this.fontSize, this.bold, this.color,
        this.rangeStart, this.rangeEnd, this.displayUnit, this.decimalPlaces,
        this.showTitle, this.title, this.titleFontSize, this.titleColor,
    ];
}

class GridlinesCard extends Card {
    name = "gridlines";
    displayName = "Gridlines";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Horizontal", value: true });
    color = new formattingSettings.ColorPicker({
        name: "color", displayName: "Color", value: { value: "#e1dfdd" }, instanceKind: CONSTANT_OR_RULE,
    });
    width = new formattingSettings.NumUpDown({ name: "width", displayName: "Width (px)", value: 1 });
    style = new formattingSettings.ItemDropdown({
        name: "style", displayName: "Style", items: LINE_STYLE_ITEMS, value: { value: "solid", displayName: "Solid" },
    });
    verticalShow = new formattingSettings.ToggleSwitch({ name: "verticalShow", displayName: "Vertical", value: false });
    verticalColor = new formattingSettings.ColorPicker({ name: "verticalColor", displayName: "Vertical color", value: { value: "#e1dfdd" } });
    verticalWidth = new formattingSettings.NumUpDown({ name: "verticalWidth", displayName: "Vertical width (px)", value: 1 });
    verticalStyle = new formattingSettings.ItemDropdown({
        name: "verticalStyle", displayName: "Vertical style", items: LINE_STYLE_ITEMS, value: { value: "solid", displayName: "Solid" },
    });

    topLevelSlice = this.show;
    slices: Slice[] = [
        this.color, this.width, this.style,
        this.verticalShow, this.verticalColor, this.verticalWidth, this.verticalStyle,
    ];
}

class ScrollbarCard extends Card {
    name = "scrollbar";
    displayName = "Scrollbar";

    // When off, the visual stops paging and fits ALL categories into the
    // viewport (thin, crammed bars) instead of showing the zoom/scroll slider.
    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show scrollbar", value: true });

    slices: Slice[] = [this.show];
}

class DataLabelsCard extends Card {
    name = "dataLabels";
    displayName = "Data labels";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", value: false });
    position = new formattingSettings.ItemDropdown({
        name: "position", displayName: "Position",
        items: [
            { value: "auto", displayName: "Auto" },
            { value: "outside", displayName: "Outside end" },
            { value: "inside", displayName: "Inside end" },
        ],
        value: { value: "auto", displayName: "Auto" },
    });
    fontFamily = new formattingSettings.FontPicker({ name: "fontFamily", displayName: "Font", value: TEXT_FONT });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Size", value: 10 });
    bold = new formattingSettings.ToggleSwitch({ name: "bold", displayName: "Bold", value: false });
    color = new formattingSettings.ColorPicker({
        name: "color", displayName: "Color", value: { value: "#252423" }, instanceKind: CONSTANT_OR_RULE,
    });
    displayUnit = new formattingSettings.ItemDropdown({
        name: "displayUnit", displayName: "Display units",
        items: DISPLAY_UNIT_ITEMS, value: { value: "auto", displayName: "Auto" },
    });
    decimalPlaces = new formattingSettings.NumUpDown({ name: "decimalPlaces", displayName: "Decimal places", value: 0 });

    topLevelSlice = this.show;
    slices: Slice[] = [
        this.position, this.fontFamily, this.fontSize, this.bold, this.color,
        this.displayUnit, this.decimalPlaces,
    ];
}

class TotalLabelsCard extends Card {
    name = "totalLabels";
    displayName = "Total labels";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", value: false });
    fontFamily = new formattingSettings.FontPicker({ name: "fontFamily", displayName: "Font", value: TEXT_FONT });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Size", value: 10 });
    bold = new formattingSettings.ToggleSwitch({ name: "bold", displayName: "Bold", value: true });
    color = new formattingSettings.ColorPicker({ name: "color", displayName: "Color", value: { value: "#252423" } });
    displayUnit = new formattingSettings.ItemDropdown({
        name: "displayUnit", displayName: "Display units",
        items: DISPLAY_UNIT_ITEMS, value: { value: "auto", displayName: "Auto" },
    });
    decimalPlaces = new formattingSettings.NumUpDown({ name: "decimalPlaces", displayName: "Decimal places", value: 0 });

    topLevelSlice = this.show;
    slices: Slice[] = [this.fontFamily, this.fontSize, this.bold, this.color, this.displayUnit, this.decimalPlaces];
}

class LegendCard extends Card {
    name = "legend";
    displayName = "Legend";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", value: true });
    position = new formattingSettings.ItemDropdown({
        name: "position", displayName: "Position",
        items: [
            { value: "top", displayName: "Top" },
            { value: "bottom", displayName: "Bottom" },
            { value: "left", displayName: "Left" },
            { value: "right", displayName: "Right" },
        ],
        value: { value: "top", displayName: "Top" },
    });
    fontFamily = new formattingSettings.FontPicker({ name: "fontFamily", displayName: "Font", value: TEXT_FONT });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Size", value: 11 });
    bold = new formattingSettings.ToggleSwitch({ name: "bold", displayName: "Bold", value: false });
    color = new formattingSettings.ColorPicker({
        name: "color", displayName: "Color", value: { value: "#605e5c" }, instanceKind: CONSTANT_OR_RULE,
    });
    titleShow = new formattingSettings.ToggleSwitch({ name: "titleShow", displayName: "Title", value: false });
    titleText = new formattingSettings.TextInput({ name: "titleText", displayName: "Title text", value: "", placeholder: "" });

    topLevelSlice = this.show;
    slices: Slice[] = [
        this.position, this.fontFamily, this.fontSize, this.bold, this.color,
        this.titleShow, this.titleText,
    ];
}

class PlotAreaCard extends Card {
    name = "plotArea";
    displayName = "Plot area";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Background", value: false });
    color = new formattingSettings.ColorPicker({ name: "color", displayName: "Color", value: { value: "#ffffff" } });
    transparency = new formattingSettings.NumUpDown({ name: "transparency", displayName: "Transparency (%)", value: 0 });

    topLevelSlice = this.show;
    slices: Slice[] = [this.color, this.transparency];
}

class TooltipCard extends Card {
    name = "tooltipSettings";
    displayName = "Tooltip";

    // Which tooltip renders on hover. "custom" = the ECharts HTML tooltip;
    // "powerbi" = the native host tooltip (report-page tooltips / drill-through).
    // Showing both at once is the double-tooltip bug this fixes.
    style = new formattingSettings.ItemDropdown({
        name: "style",
        displayName: "Style",
        items: [
            { value: "custom", displayName: "Custom" },
            { value: "powerbi", displayName: "Power BI (native)" },
        ],
        value: { value: "custom", displayName: "Custom" },
    });

    slices: Slice[] = [this.style];
}

class ConstantLineCard extends Card {
    name = "constantLine";
    displayName = "Constant line";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show", value: false });
    value = new formattingSettings.NumUpDown({ name: "value", displayName: "Value", value: 0 });
    color = new formattingSettings.ColorPicker({ name: "color", displayName: "Color", value: { value: "#e44c4c" } });
    lineStyle = new formattingSettings.ItemDropdown({
        name: "lineStyle", displayName: "Style", items: LINE_STYLE_ITEMS, value: { value: "dashed", displayName: "Dashed" },
    });
    lineWidth = new formattingSettings.NumUpDown({ name: "lineWidth", displayName: "Width (px)", value: 2 });
    showLabel = new formattingSettings.ToggleSwitch({ name: "showLabel", displayName: "Label", value: false });
    labelText = new formattingSettings.TextInput({ name: "labelText", displayName: "Label text", value: "", placeholder: "" });
    labelColor = new formattingSettings.ColorPicker({ name: "labelColor", displayName: "Label color", value: { value: "#e44c4c" } });

    topLevelSlice = this.show;
    slices: Slice[] = [this.value, this.color, this.lineStyle, this.lineWidth, this.showLabel, this.labelText, this.labelColor];
}

export class EChartsColumnChartFormattingSettings extends Model {
    dataColors = new DataColorsCard();
    columns = new ColumnsCard();
    xAxis = new XAxisCard();
    yAxis = new YAxisCard();
    scrollbar = new ScrollbarCard();
    gridlines = new GridlinesCard();
    dataLabels = new DataLabelsCard();
    totalLabels = new TotalLabelsCard();
    legend = new LegendCard();
    plotArea = new PlotAreaCard();
    constantLine = new ConstantLineCard();
    tooltip = new TooltipCard();

    cards = [
        this.dataColors, this.columns, this.xAxis, this.yAxis, this.scrollbar,
        this.gridlines, this.dataLabels, this.totalLabels, this.legend,
        this.plotArea, this.constantLine, this.tooltip,
    ];
}

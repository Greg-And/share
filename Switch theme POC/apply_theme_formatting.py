#!/usr/bin/env python3
"""
apply_theme_formatting_v12.py
-------------------------------------------------------------------------------
Bulk-apply dynamic theme conditional formatting to a Power BI PBIP report saved
in the ENHANCED report format (PBIR).

Tuned against the SECOND revision of "switch theme POC.Report", where the author
had already turned several things off by hand.

v12 changes vs v11 - all driven by what the updated report contains:

  * NEVER writes 'show'. v11 forced background.show = true and border.show =
    true. In the updated report the two cards have background.show = false and
    border.show = false on purpose, and the shape has outline.show = false.
    Re-running v11 would have switched all of those back on. v12 writes colours
    only, so every visibility decision in the file survives untouched.

  * SERIES COLOURS. lineClusteredColumnComboChart carries Sales on role 'Y'
    (the columns) and Profit% on role 'Y2' (the line), so they now bind to
    clr_ColumnSeries and clr_LineSeries respectively. waterfallChart has no
    line - its bars are semantic, so sentimentColors gets clr_Good / clr_Bad /
    clr_Neutral. Plain column and bar charts get clr_ColumnSeries, plain line
    and area charts get clr_LineSeries.

  * EXISTING dataPoint BINDINGS ARE LEFT ALONE. The donut is already bound to a
    hand-written measure called 'Country color slices'. Overwriting that would
    destroy the author's work, so a role that already has a fill is skipped
    unless OVERWRITE_SERIES_COLORS is turned on.

  * shape fill confirmed as fill.fillColor - the updated file shows Desktop
    wrote exactly that, so the guess in v11 was right and the config pair is
    kept only for other reports.

MEASURES expected in the model (table set by MEASURE_TABLE):
    clr_PageBackground  clr_VisualBackground  clr_VisualBorder
    clr_VisualHeaderBackground  clr_FontPrimary  clr_FontSecondary
    clr_TableGridline  clr_ChartGridline  clr_BorderSubtle
    clr_ColumnSeries  clr_LineSeries  clr_Good  clr_Bad  clr_Neutral

RUNNING FROM VISUAL STUDIO
    Set REPORT_PATH below, press F5. Starts in DRY_RUN.
"""

import json
import os
import sys
import shutil
import datetime

# =============================================================================
# CONFIG
# =============================================================================

REPORT_PATH = r""
DRY_RUN = True

MEASURE_TABLE = "Theme"

M_PAGE_BG      = "clr_PageBackground"
M_VISUAL_BG    = "clr_VisualBackground"
M_VISUAL_BORD  = "clr_VisualBorder"
M_HEADER_BG    = "clr_VisualHeaderBackground"
M_FONT_PRIMARY = "clr_FontPrimary"
M_FONT_SECOND  = "clr_FontSecondary"
M_GRIDLINE     = "clr_TableGridline"
M_CHART_GRID   = "clr_ChartGridline"
M_BORDER_SUB   = "clr_BorderSubtle"

M_COLUMN       = "clr_ColumnSeries"
M_LINE         = "clr_LineSeries"
M_GOOD         = "clr_Good"
M_BAD          = "clr_Bad"
M_NEUTRAL      = "clr_Neutral"

BACKGROUND_SHAPE_NAME = "Report background"
SHAPE_FILL_OBJECT   = "fill"
SHAPE_FILL_PROPERTY = "fillColor"

HEADER_FONT = M_FONT_SECOND
SET_CHART_GRIDLINES = True
SET_PAGE_CANVAS = False

# Bind dataPoint / sentimentColors on charts.
SET_SERIES_COLORS = True

# A role that already carries a fill is left alone, so hand-written bindings
# such as the donut's 'Country color slices' survive. Turn on only if you want
# this script to become the single source of truth for series colours.
OVERWRITE_SERIES_COLORS = False

# Waterfall bars are increase / decrease / total. True maps them to the
# semantic measures; False paints every bar with clr_ColumnSeries.
WATERFALL_USE_SENTIMENT = True

LINE_ENDING = "\r\n"
JSON_INDENT = 2

# =============================================================================
# Visual type buckets
# =============================================================================

TABLE_TYPES  = {"tableEx"}
MATRIX_TYPES = {"pivotTable"}

COLUMN_BAR_CHARTS = {
    "columnChart", "clusteredColumnChart", "stackedColumnChart",
    "hundredPercentStackedColumnChart", "barChart", "clusteredBarChart",
    "stackedBarChart", "hundredPercentStackedBarChart",
}
LINE_AREA_CHARTS = {"lineChart", "areaChart", "stackedAreaChart"}
COMBO_CHARTS     = {"lineStackedColumnComboChart", "lineClusteredColumnComboChart"}
WATERFALL_CHARTS = {"waterfallChart"}

CHART_TYPES = COLUMN_BAR_CHARTS | LINE_AREA_CHARTS | COMBO_CHARTS | WATERFALL_CHARTS | {
    "scatterChart", "pieChart", "donutChart", "ribbonChart", "funnel", "treemap",
    "gauge", "kpi", "filledMap", "map", "decompositionTreeVisual", "keyDriversVisual",
}

AXIS_LESS_CHARTS = {"pieChart", "donutChart", "treemap", "funnel", "gauge",
                    "kpi", "filledMap", "map", "decompositionTreeVisual",
                    "keyDriversVisual"}

MODERN_CARD_TYPES   = {"cardVisual"}
LEGACY_CARD_TYPES   = {"card", "multiRowCard"}
MODERN_SLICER_TYPES = {"advancedSlicerVisual", "listSlicer", "textSlicer"}
LEGACY_SLICER_TYPES = {"slicer"}
BUTTON_TYPES        = {"actionButton", "pageNavigator", "bookmarkNavigator"}

DECORATIVE_TYPES = {
    "image", "shape", "basicShape", "textbox",
    "actionButton", "pageNavigator", "bookmarkNavigator",
}

# =============================================================================
# expr helpers
# =============================================================================

def measure_color(measure_name):
    return {"solid": {"color": {"expr": {"Measure": {
        "Expression": {"SourceRef": {"Entity": MEASURE_TABLE}},
        "Property": measure_name,
    }}}}}


def set_prop(objects, obj_name, props, only_if_exists=False):
    """Merge props into objects[obj_name][0]['properties'].

    Existing keys, selectors and sibling entries are preserved. Nothing here
    ever writes 'show' - see the module docstring.
    """
    arr = objects.get(obj_name)
    if only_if_exists and (not isinstance(arr, list) or not arr):
        return False
    if not isinstance(arr, list) or not arr:
        arr = [{"properties": {}}]
    if not isinstance(arr[0].get("properties"), dict):
        arr[0]["properties"] = {}
    arr[0]["properties"].update(props)
    objects[obj_name] = arr
    return True


def set_font(objects, obj_name, keys, measure=M_FONT_SECOND):
    set_prop(objects, obj_name, {k: measure_color(measure) for k in keys})


def set_data_labels(objects, measure=M_FONT_SECOND):
    """Data labels need a default entry AND a dataViewWildcard entry, otherwise
    the colour is evaluated but never painted on some visual types."""
    arr = objects.get("labels")
    if not isinstance(arr, list) or not arr:
        arr = [{"properties": {}}]
    if not isinstance(arr[0].get("properties"), dict):
        arr[0]["properties"] = {}
    arr[0]["properties"]["color"] = measure_color(measure)

    def _is_wildcard(entry):
        try:
            return any("dataViewWildcard" in d for d in entry["selector"]["data"])
        except Exception:
            return False

    if not any(_is_wildcard(e) for e in arr if isinstance(e, dict)):
        arr.append({
            "properties": {"color": measure_color(measure)},
            "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]},
        })
    else:
        for e in arr:
            if isinstance(e, dict) and _is_wildcard(e):
                e.setdefault("properties", {})["color"] = measure_color(measure)
    objects["labels"] = arr

# =============================================================================
# Series colours
# =============================================================================

def _entry_roles(entry):
    try:
        roles = []
        for d in entry["selector"]["data"]:
            roles.extend(d.get("roles") or [])
        return roles
    except Exception:
        return []


def set_datapoint_role(objects, role, measure):
    """Bind dataPoint.fill for one query role, e.g. 'Y' or 'Y2'.

    Mirrors the selector shape Desktop itself writes:
        {"properties": {"fill": ...}, "selector": {"data": [{"roles": ["Y"]}]}}

    A role that already has a fill is skipped unless OVERWRITE_SERIES_COLORS,
    so hand-written bindings are never destroyed.
    """
    arr = objects.get("dataPoint")
    if not isinstance(arr, list):
        arr = []

    for entry in arr:
        if not isinstance(entry, dict):
            continue
        if role in _entry_roles(entry):
            has_fill = "fill" in (entry.get("properties") or {})
            if has_fill and not OVERWRITE_SERIES_COLORS:
                return "skipped"
            entry.setdefault("properties", {})["fill"] = measure_color(measure)
            objects["dataPoint"] = arr
            return "updated"

    # A bare entry with no selector would repaint every series, so only add a
    # role-scoped one.
    arr.append({
        "properties": {"fill": measure_color(measure)},
        "selector": {"data": [{"roles": [role]}]},
    })
    objects["dataPoint"] = arr
    return "added"


def has_any_datapoint_fill(objects):
    return any(
        isinstance(e, dict) and "fill" in (e.get("properties") or {})
        for e in (objects.get("dataPoint") or [])
    )


def set_series_colors(objects, vtype, log):
    if not SET_SERIES_COLORS:
        return

    if vtype in COMBO_CHARTS:
        # Y = the clustered / stacked columns, Y2 = the line.
        log.append(("Y -> " + M_COLUMN, set_datapoint_role(objects, "Y", M_COLUMN)))
        log.append(("Y2 -> " + M_LINE, set_datapoint_role(objects, "Y2", M_LINE)))

    elif vtype in WATERFALL_CHARTS:
        if WATERFALL_USE_SENTIMENT:
            props = {
                "increaseFill": measure_color(M_GOOD),
                "decreaseFill": measure_color(M_BAD),
                "totalFill": measure_color(M_NEUTRAL),
                "otherFill": measure_color(M_COLUMN),
            }
        else:
            c = measure_color(M_COLUMN)
            props = {"increaseFill": c, "decreaseFill": c,
                     "totalFill": c, "otherFill": c}
        existing = objects.get("sentimentColors")
        already = bool(existing) and "increaseFill" in (existing[0].get("properties") or {})
        if already and not OVERWRITE_SERIES_COLORS:
            log.append(("sentimentColors", "skipped"))
        else:
            set_prop(objects, "sentimentColors", props)
            log.append(("sentimentColors", "updated" if already else "added"))

    elif vtype in COLUMN_BAR_CHARTS:
        log.append(("Y -> " + M_COLUMN, set_datapoint_role(objects, "Y", M_COLUMN)))

    elif vtype in LINE_AREA_CHARTS:
        log.append(("Y -> " + M_LINE, set_datapoint_role(objects, "Y", M_LINE)))

# =============================================================================
# Identifying the background rectangle
# =============================================================================

def _literal_text(container, obj_name, prop):
    try:
        raw = container[obj_name][0]["properties"][prop]["expr"]["Literal"]["Value"]
    except Exception:
        return None
    return raw.strip().strip("'") if isinstance(raw, str) else None


def visual_display_name(vjson):
    visual = vjson.get("visual") or {}
    container = visual.get("visualContainerObjects") or {}
    for c in (
        _literal_text(container, "title", "text"),
        vjson.get("displayName"),
        visual.get("displayName"),
        _literal_text(visual.get("objects") or {}, "general", "altText"),
        vjson.get("name"),
    ):
        if isinstance(c, str) and c.strip():
            return c.strip()
    return ""


def is_background_shape(vjson):
    return visual_display_name(vjson).casefold() == BACKGROUND_SHAPE_NAME.casefold()

# =============================================================================
# Formatting
# =============================================================================

def format_background_shape(vjson):
    """Page-background rectangle: fill colour only. The outline is left exactly
    as the author set it, including outline.show = false."""
    visual = vjson["visual"]
    objects = visual.setdefault("objects", {})
    set_prop(objects, SHAPE_FILL_OBJECT,
             {SHAPE_FILL_PROPERTY: measure_color(M_PAGE_BG)})
    container = visual.setdefault("visualContainerObjects", {})
    set_prop(container, "background",
             {"color": measure_color(M_PAGE_BG)}, only_if_exists=True)
    return "background-shape"


def format_chrome(visual, decorative):
    """Container chrome. Colours only - visibility is never touched, so a
    background or border switched off by hand stays off."""
    container = visual.setdefault("visualContainerObjects", {})

    if decorative:
        set_prop(container, "background",
                 {"color": measure_color(M_VISUAL_BG)}, only_if_exists=True)
        set_prop(container, "border",
                 {"color": measure_color(M_VISUAL_BORD)}, only_if_exists=True)
    else:
        set_prop(container, "background", {"color": measure_color(M_VISUAL_BG)})
        set_prop(container, "border", {"color": measure_color(M_VISUAL_BORD)})

    set_prop(container, "title", {
        "fontColor": measure_color(M_FONT_PRIMARY),
        "background": measure_color(M_HEADER_BG),
    })
    set_prop(container, "subTitle",
             {"fontColor": measure_color(M_FONT_SECOND)}, only_if_exists=True)
    set_prop(container, "divider",
             {"color": measure_color(M_VISUAL_BORD)}, only_if_exists=True)

    if decorative:
        return

    set_prop(container, "visualHeader", {
        "background": measure_color(M_VISUAL_BG),
        "border": measure_color(M_VISUAL_BORD),
        "foreground": measure_color(M_FONT_SECOND),
    })

    # No documented, measure-bindable border exists on the modern tooltip, so
    # background and fonts only.
    set_prop(container, "visualTooltip", {
        "titleFontColor": measure_color(M_FONT_PRIMARY),
        "themedTitleFontColor": measure_color(M_FONT_PRIMARY),
        "valueFontColor": measure_color(M_FONT_SECOND),
        "themedValueFontColor": measure_color(M_FONT_SECOND),
        "actionFontColor": measure_color(M_FONT_SECOND),
        "background": measure_color(M_VISUAL_BG),
        "themedBackground": measure_color(M_VISUAL_BG),
    })


def format_table_like(objects, is_matrix):
    set_prop(objects, "values", {
        "fontColorPrimary": measure_color(M_FONT_SECOND),
        "backColorPrimary": measure_color(M_VISUAL_BG),
        "fontColorSecondary": measure_color(M_FONT_SECOND),
        "backColorSecondary": measure_color(M_VISUAL_BG),
    })
    set_prop(objects, "columnHeaders", {
        "fontColor": measure_color(HEADER_FONT),
        "backColor": measure_color(M_HEADER_BG),
    })
    set_prop(objects, "grid", {
        "gridVerticalColor": measure_color(M_GRIDLINE),
        "gridHorizontalColor": measure_color(M_GRIDLINE),
        "outlineColor": measure_color(M_BORDER_SUB),
    })
    set_prop(objects, "total", {
        "fontColor": measure_color(M_FONT_SECOND),
        "backColor": measure_color(M_VISUAL_BG),
    })
    # 'total' and 'subTotals' drive the Row/Column SUBTOTAL sections. The matrix
    # grand totals are separate objects with no fx button in the UI - Desktop
    # writes a ThemeDataColor there, which we swap for a measure binding.
    # only_if_exists: the objects appear only once someone has set a grand-total
    # colour by hand, and creating them from scratch would be a guess.
    set_prop(objects, "columnTotal",
             {"fontColor": measure_color(M_FONT_SECOND)}, only_if_exists=True)
    set_prop(objects, "rowTotal",
             {"fontColor": measure_color(M_FONT_SECOND)}, only_if_exists=True)
    if is_matrix:
        set_prop(objects, "rowHeaders", {
            "fontColor": measure_color(HEADER_FONT),
            "backColor": measure_color(M_HEADER_BG),
        })
        set_prop(objects, "subTotals", {
            "fontColor": measure_color(M_FONT_SECOND),
            "backColor": measure_color(M_VISUAL_BG),
        })


def format_chart(objects, vtype, log):
    if vtype not in AXIS_LESS_CHARTS:
        axis = {"labelColor": measure_color(M_FONT_SECOND),
                "titleColor": measure_color(M_FONT_SECOND)}
        if SET_CHART_GRIDLINES:
            axis["gridlineColor"] = measure_color(M_CHART_GRID)
        set_prop(objects, "categoryAxis", dict(axis))
        set_prop(objects, "valueAxis", dict(axis))

    set_font(objects, "legend", ["labelColor", "titleColor"])
    set_data_labels(objects)

    if vtype in {"pieChart", "donutChart", "funnel"}:
        set_font(objects, "percentLabels", ["color"])
    if vtype in {"donutChart", "gauge", "kpi"}:
        set_font(objects, "calloutValue", ["fontColor", "color"])
    if vtype in {"gauge", "kpi"}:
        set_font(objects, "indicator", ["fontColor", "color"])
        set_font(objects, "goals", ["color"])
    if vtype in {"treemap", "filledMap", "map", "decompositionTreeVisual"}:
        set_font(objects, "categoryLabels", ["color", "fontColor"])

    set_series_colors(objects, vtype, log)


def format_modern_card(objects):
    set_font(objects, "value", ["fontColor"])
    set_font(objects, "label", ["fontColor"])
    set_font(objects, "referenceLabelTitle", ["titleFontColor"])
    set_font(objects, "referenceLabelValue", ["valueFontColor"])
    set_font(objects, "referenceLabelDetail", ["detailFontColor"])
    set_prop(objects, "fillCustom",
             {"fillColor": measure_color(M_VISUAL_BG)}, only_if_exists=True)
    set_prop(objects, "referenceLabel",
             {"backgroundColor": measure_color(M_VISUAL_BG)}, only_if_exists=True)
    set_prop(objects, "divider",
             {"dividerColor": measure_color(M_BORDER_SUB)}, only_if_exists=True)
    set_prop(objects, "outline",
             {"lineColor": measure_color(M_VISUAL_BORD)}, only_if_exists=True)
    set_prop(objects, "borderCustom",
             {"borderColor": measure_color(M_VISUAL_BORD)}, only_if_exists=True)


def format_modern_slicer(objects):
    """The slicer already carries per-state selectors written by Desktop
    (default vs selection:selected). Only the plain entries are recoloured, so
    the selected-state binding to clr_ButtonFillSelected survives."""
    set_font(objects, "value", ["fontColor"], M_FONT_PRIMARY)
    set_font(objects, "label", ["fontColor"], M_FONT_PRIMARY)
    set_prop(objects, "outline",
             {"lineColor": measure_color(M_BORDER_SUB)}, only_if_exists=True)


def format_content(visual, vtype, decorative, log):
    objects = visual.setdefault("objects", {})

    if vtype in TABLE_TYPES or vtype in MATRIX_TYPES:
        format_table_like(objects, vtype in MATRIX_TYPES)
    elif vtype in CHART_TYPES:
        format_chart(objects, vtype, log)
    elif vtype in MODERN_CARD_TYPES:
        format_modern_card(objects)
    elif vtype in LEGACY_CARD_TYPES:
        set_font(objects, "calloutValue", ["fontColor", "color"])
        set_font(objects, "categoryLabels", ["color", "fontColor"])
        set_font(objects, "dataLabels", ["color", "fontColor"])
        set_data_labels(objects)
    elif vtype in MODERN_SLICER_TYPES:
        format_modern_slicer(objects)
    elif vtype in LEGACY_SLICER_TYPES:
        set_font(objects, "header", ["fontColor"])
        set_font(objects, "items", ["fontColor"])
        set_prop(objects, "general", {"outlineColor": measure_color(M_VISUAL_BORD)})
    elif vtype in BUTTON_TYPES:
        set_font(objects, "text", ["fontColor"])
        set_font(objects, "icon", ["lineColor"])
    elif vtype == "textbox":
        pass
    elif not decorative:
        set_data_labels(objects)


def format_visual(vjson, log):
    visual = vjson.get("visual")
    if not isinstance(visual, dict):
        return None
    if is_background_shape(vjson):
        return format_background_shape(vjson)
    vtype = visual.get("visualType", "")
    decorative = vtype in DECORATIVE_TYPES
    format_chrome(visual, decorative)
    format_content(visual, vtype, decorative, log)
    return vtype or "unknown"

# =============================================================================
# I/O
# =============================================================================

def write_json(path, data):
    text = json.dumps(data, indent=JSON_INDENT, ensure_ascii=False)
    with open(path, "w", encoding="utf-8", newline=LINE_ENDING) as f:
        f.write(text)


def format_page(page_json_path):
    if not SET_PAGE_CANVAS:
        return False
    with open(page_json_path, "r", encoding="utf-8") as f:
        pj = json.load(f)
    set_prop(pj.setdefault("objects", {}), "background",
             {"color": measure_color(M_PAGE_BG)})
    if not DRY_RUN:
        write_json(page_json_path, pj)
    return True

# =============================================================================
# Main
# =============================================================================

def resolve_report_path():
    if len(sys.argv) > 1:
        return sys.argv[1]
    if REPORT_PATH.strip():
        return REPORT_PATH.strip()
    print("ERROR: no report path. Set REPORT_PATH at the top of this file, or\n"
          '       run: python apply_theme_formatting_v12.py "C:/path/Your.Report"')
    sys.exit(1)


def main():
    report_root = resolve_report_path()
    pages_dir = os.path.join(report_root, "definition", "pages")
    if not os.path.isdir(pages_dir):
        print(f"ERROR: {pages_dir} not found.")
        print("       Point me at the *.Report folder of a PBIR-format report.")
        sys.exit(1)

    if DRY_RUN:
        print("DRY RUN - nothing will be written.\n")
    else:
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = report_root.rstrip("/\\") + f"_backup_{stamp}"
        shutil.copytree(report_root, backup)
        print(f"Backup created: {backup}\n")

    counts, pages_done, bg_found = {}, 0, 0

    for page_name in sorted(os.listdir(pages_dir)):
        page_folder = os.path.join(pages_dir, page_name)
        if not os.path.isdir(page_folder):
            continue
        pages_done += 1

        page_json = os.path.join(page_folder, "page.json")
        if os.path.isfile(page_json):
            format_page(page_json)

        visuals_folder = os.path.join(page_folder, "visuals")
        if not os.path.isdir(visuals_folder):
            continue

        for vname in sorted(os.listdir(visuals_folder)):
            vpath = os.path.join(visuals_folder, vname, "visual.json")
            if not os.path.isfile(vpath):
                continue
            with open(vpath, "r", encoding="utf-8") as f:
                vjson = json.load(f)

            log = []
            tag = format_visual(vjson, log)
            if tag is None:
                continue
            if tag == "background-shape":
                bg_found += 1
                print(f"  {BACKGROUND_SHAPE_NAME}: "
                      f"{SHAPE_FILL_OBJECT}.{SHAPE_FILL_PROPERTY} = {M_PAGE_BG}")
            for what, how in log:
                print(f"  {tag}: {what}  [{how}]")
            counts[tag] = counts.get(tag, 0) + 1

            if not DRY_RUN:
                write_json(vpath, vjson)

    print(f"\nProcessed {sum(counts.values())} visuals across {pages_done} pages.")
    for tag in sorted(counts):
        print(f"  {counts[tag]:4d}  {tag}")

    if bg_found == 0:
        print(f"\nWARNING: no visual named '{BACKGROUND_SHAPE_NAME}' found.")

    print(f"\nMeasures bound from table: '{MEASURE_TABLE}'")
    print("No 'show' property was written - all visibility settings preserved.")
    if DRY_RUN:
        print("\nDRY RUN complete - set DRY_RUN = False to write.")


if __name__ == "__main__":
    main()
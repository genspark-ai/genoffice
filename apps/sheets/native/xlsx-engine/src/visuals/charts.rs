//! `chartN.xml` reader: chart kind, series, axes, legend and data labels,
//! plus the cached value / category extraction the renderer plots from.

use super::*;

pub(crate) const CHART_TYPE_NAMES: [&str; 7] = [
    "barChart",
    "lineChart",
    "pieChart",
    "doughnutChart",
    "areaChart",
    "scatterChart",
    "radarChart",
];

/// 3D plot elements fold onto their flat pipeline at parse time (flat 2D
/// projection of the right chart type); the wire only carries the 2D name.
/// stockChart/surfaceChart stay unmapped.
pub(crate) const CHART_TYPE_3D_FOLDS: [(&str, &str); 4] = [
    ("bar3DChart", "barChart"),
    ("line3DChart", "lineChart"),
    ("pie3DChart", "pieChart"),
    ("area3DChart", "areaChart"),
];

/// The 2D wire name of a plot element, folding 3D variants.
pub(crate) fn flat_plot_name(node: Node<'_, '_>) -> Option<&'static str> {
    let name = node.tag_name().name();
    CHART_TYPE_NAMES
        .iter()
        .copied()
        .find(|flat| *flat == name)
        .or_else(|| {
            CHART_TYPE_3D_FOLDS
                .iter()
                .find(|(three_d, _)| *three_d == name)
                .map(|(_, flat)| *flat)
        })
}

pub(crate) fn read_chart(
    archive: &mut ZipArchive<File>,
    chart_path: &str,
    colors: &ColorContext,
    formats: &mut SourceFormats,
) -> Result<ChartMetadata, SidecarError> {
    let xml = read_xml(archive, chart_path)?;
    let document = parse_document(&xml, chart_path)?;
    Ok(chart_metadata(&document, colors, &mut |reference| {
        formats.first_cell_format(archive, reference)
    }))
}

/// Number format of the first cell behind a `c:f` reference (None when the
/// cells are General/text or the reference cannot be resolved).
pub(crate) type SourceFormatLookup<'a> = dyn FnMut(&str) -> Option<String> + 'a;

/// Format candidates behind a `sourceLinked="1"` numFmt, taken from the
/// first series the way Excel does: its source cells, then its numCache.
#[derive(Default)]
pub(crate) struct LinkedFormats {
    value_source: Option<String>,
    value_cache: Option<String>,
    category_source: Option<String>,
    category_cache: Option<String>,
}

impl LinkedFormats {
    fn from_first_series(
        series: Option<Node<'_, '_>>,
        lookup: &mut SourceFormatLookup<'_>,
    ) -> Self {
        let Some(series) = series else {
            return Self::default();
        };
        let value_node = direct_child(series, "val").or_else(|| direct_child(series, "yVal"));
        let category_node = direct_child(series, "cat").or_else(|| direct_child(series, "xVal"));
        Self {
            value_source: value_node
                .and_then(formula_ref)
                .and_then(|reference| lookup(&reference)),
            value_cache: value_node.and_then(cache_format_code),
            category_source: category_node
                .filter(is_numeric_ref)
                .and_then(formula_ref)
                .and_then(|reference| lookup(&reference)),
            category_cache: category_node.and_then(cache_format_code),
        }
    }

    /// The attribute's formatCode unless it is source-linked, in which case
    /// the source cells win, then the cache, then the (stale) literal.
    fn resolve(&self, num_fmt: Option<Node<'_, '_>>, category: bool) -> Option<String> {
        let node = num_fmt?;
        let literal = node.attribute("formatCode").map(ToOwned::to_owned);
        if !matches!(node.attribute("sourceLinked"), Some("1") | Some("true")) {
            return literal;
        }
        let (source, cache) = if category {
            (&self.category_source, &self.category_cache)
        } else {
            (&self.value_source, &self.value_cache)
        };
        source.clone().or_else(|| cache.clone()).or(literal)
    }
}

/// The series a plot axis scales: the first c:ser of the plot group that
/// lists the axis's c:axId (a combo's secondary axis has its own group).
fn axis_series<'a>(document: &'a Document<'a>, axis: Node<'a, 'a>) -> Option<Node<'a, 'a>> {
    let id = direct_child(axis, "axId")?.attribute("val")?;
    document
        .descendants()
        .filter(|node| flat_plot_name(*node).is_some())
        .find(|plot| {
            plot.children()
                .any(|child| child.has_tag_name("axId") && child.attribute("val") == Some(id))
        })
        .and_then(|plot| plot.children().find(|child| child.has_tag_name("ser")))
}

/// strRef categories are text: their cells' number format does not apply.
fn is_numeric_ref(node: &Node<'_, '_>) -> bool {
    node.descendants().any(|child| child.has_tag_name("numRef"))
}

/// Every `c:f` reference in a chart part, so the source-format lookup can be
/// warmed in one worksheet pass instead of one prefix scan per reference.
pub(crate) fn chart_formula_references(
    archive: &mut ZipArchive<File>,
    chart_path: &str,
) -> Vec<String> {
    let Ok(xml) = read_xml(archive, chart_path) else {
        return Vec::new();
    };
    let Ok(document) = parse_document(&xml, chart_path) else {
        return Vec::new();
    };
    document
        .descendants()
        .filter(|node| node.is_element() && node.has_tag_name("f"))
        .filter_map(|node| node.text())
        .filter(|value| !value.is_empty())
        .map(ToOwned::to_owned)
        .collect()
}

/// Font shorthand of a node's c:txPr//a:defRPr: size, bold, solid fill.
pub(crate) fn text_style(parent: Node<'_, '_>, colors: &ColorContext) -> Option<ChartTitleStyle> {
    let def = direct_child(parent, "txPr")?
        .descendants()
        .find(|child| child.has_tag_name("defRPr"))?;
    Some(ChartTitleStyle {
        size: def
            .attribute("sz")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value / 100.0),
        bold: def
            .attribute("b")
            .map(|value| value == "1" || value == "true"),
        color: drawing_fill_color(def, colors),
    })
}

pub(crate) fn chart_metadata(
    document: &Document<'_>,
    colors: &ColorContext,
    lookup: &mut SourceFormatLookup<'_>,
) -> ChartMetadata {
    let chart_types = CHART_TYPE_NAMES
        .iter()
        .filter(|name| {
            document
                .descendants()
                .any(|node| flat_plot_name(node) == Some(**name))
        })
        .map(|name| (*name).to_owned())
        .collect::<Vec<_>>();
    // Only the chart-level title — axes carry their own c:title deeper down.
    let chart_node = document
        .descendants()
        .find(|node| node.has_tag_name("chart"));
    let title_node = chart_node.and_then(|chart| direct_child(chart, "title"));
    let explicit_title = title_node
        .map(|node| {
            let rich = node
                .descendants()
                .filter(|child| child.has_tag_name("t"))
                .filter_map(|child| child.text())
                .collect::<String>();
            if !rich.is_empty() {
                return rich;
            }
            // Cell-linked title (<c:tx><c:strRef>): the strCache <c:v> holds
            // the cached cell text — show that instead of a placeholder (#181)
            node.descendants()
                .filter(|child| child.has_tag_name("v"))
                .filter_map(|child| child.text())
                .collect::<String>()
        })
        .filter(|value| !value.is_empty());
    let auto_title_deleted = chart_node
        .and_then(|chart| direct_child(chart, "autoTitleDeleted"))
        .and_then(|node| node.attribute("val"))
        .is_some_and(|value| value == "1" || value == "true");
    let bar_direction = document
        .descendants()
        .find(|node| node.has_tag_name("barDir"))
        .and_then(|node| node.attribute("val"))
        .map(ToOwned::to_owned);
    let series_nodes = document
        .descendants()
        .filter(|node| node.has_tag_name("ser"))
        .collect::<Vec<_>>();
    let sole_series_named = matches!(&series_nodes[..],
        [only] if direct_child(*only, "tx").and_then(first_cached_value).is_some());
    let linked = LinkedFormats::from_first_series(series_nodes.first().copied(), lookup);
    let mut series = series_nodes
        .iter()
        .enumerate()
        .map(|(index, node)| parse_chart_series(*node, index, colors, lookup))
        .collect::<Vec<_>>();
    trim_blank_tail(&mut series);
    // Excel's title rules: explicit text wins; a deleted auto title (or no
    // <c:title> at all) means no title; a present-but-empty <c:title> shows
    // the auto title — the sole series' name, else the "Chart Title"
    // placeholder. An empty string means "no title" on the wire.
    let title = match explicit_title {
        Some(text) => text,
        None if auto_title_deleted || title_node.is_none() => String::new(),
        None => match &series[..] {
            [only] if sole_series_named => only.name.clone(),
            _ => "Chart Title".into(),
        },
    };
    let (x_axis, y_axis, secondary_y_axis) = axis_infos(
        document,
        colors,
        &linked,
        series_nodes.first().copied(),
        lookup,
    );
    let area_fill = |node: Option<Node<'_, '_>>| {
        node.and_then(|node| direct_child(node, "spPr"))
            .and_then(|sppr| area_fill_color(sppr, colors))
    };
    let chart_area_fill = area_fill(Some(document.root_element()));
    let plot_area_fill = area_fill(chart_node.and_then(|chart| direct_child(chart, "plotArea")));
    let title_style = title_node.and_then(|node| text_style(node, colors));
    let data_label_style = data_labels_node(document).and_then(|node| text_style(node, colors));
    ChartMetadata {
        chart_types,
        bar_direction,
        title,
        legend: legend_position(document),
        data_labels: data_labels(document),
        data_label_position: data_label_position(document),
        data_label_format: data_label_format(document, &linked),
        axis_titles: axis_titles(document),
        grouping: plot_grouping(document),
        gridlines: value_axis(document).map(|axis| direct_child(axis, "majorGridlines").is_some()),
        value_axis: value_axis_bounds(document),
        category_axis_format: category_axis_format(document, &linked),
        gap_width_pct: plot_val_attribute(document, "barChart", "gapWidth"),
        hole_size_pct: plot_val_attribute(document, "doughnutChart", "holeSize"),
        x_axis,
        y_axis,
        scatter_style: document
            .descendants()
            .find(|node| node.has_tag_name("scatterStyle"))
            .and_then(|node| node.attribute("val"))
            .map(ToOwned::to_owned),
        // Direct child only: c:ser/c:marker (a symbol container) must not
        // register as the plot flag. CT_Boolean: a bare <c:marker/> is true.
        line_markers: document
            .descendants()
            .find(|node| node.has_tag_name("lineChart") || node.has_tag_name("line3DChart"))
            .and_then(|plot| direct_child(plot, "marker"))
            .map(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
        secondary_y_axis,
        disp_blanks_as: chart_node
            .and_then(|chart| direct_child(chart, "dispBlanksAs"))
            .and_then(|node| node.attribute("val"))
            .filter(|value| matches!(*value, "gap" | "zero" | "span"))
            .map(ToOwned::to_owned),
        title_style,
        chart_area_fill,
        plot_area_fill,
        data_label_style,
        series,
    }
}

/// Excel ends the plot at the last populated point: a trailing run whose
/// categories are blank and whose cached values are all blank or zero (a
/// template's pre-sized empty rows) gets no slots, no axis labels.
pub(crate) fn trim_blank_tail(series_list: &mut [ChartSeries]) {
    let has_categories = series_list.iter().any(|series| {
        series
            .categories
            .iter()
            .any(|label| !label.trim().is_empty())
    });
    if !has_categories {
        return;
    }
    let len = series_list
        .iter()
        .map(|series| series.values.len().max(series.categories.len()))
        .max()
        .unwrap_or(0);
    let occupied = |index: usize| {
        series_list.iter().any(|series| {
            series
                .categories
                .get(index)
                .is_some_and(|label| !label.trim().is_empty())
                // Blank cache slots hold a 0 filler, so nonzero means data.
                || series.values.get(index).is_some_and(|value| *value != 0.0)
        })
    };
    let keep = (0..len)
        .rev()
        .find(|index| occupied(*index))
        .map_or(1, |index| index + 1);
    if keep >= len {
        return;
    }
    for series in series_list {
        series.categories.truncate(keep);
        series.values.truncate(keep);
        if let Some(blanks) = series.blanks.take() {
            let kept: Vec<usize> = blanks.into_iter().filter(|index| *index < keep).collect();
            series.blanks = (!kept.is_empty()).then_some(kept);
        }
        if let Some(groups) = series.category_groups.take() {
            let kept: Vec<CategoryGroup> = groups
                .into_iter()
                .filter_map(|mut group| {
                    group.end = group.end.min(keep);
                    (group.start < group.end).then_some(group)
                })
                .collect();
            series.category_groups = (!kept.is_empty()).then_some(kept);
        }
    }
}

/// All plot axes keyed by side: axPos b/t → X, l/r → Y. Falls back to the
/// element kind (catAx → X, valAx → Y) when axPos is missing.
pub(crate) fn axis_infos<'a>(
    document: &'a Document<'a>,
    colors: &ColorContext,
    linked: &LinkedFormats,
    first_series: Option<Node<'a, 'a>>,
    lookup: &mut SourceFormatLookup<'_>,
) -> (Option<AxisInfo>, Option<AxisInfo>, Option<AxisInfo>) {
    // Scatter plots have two valAx: the bottom one scales xVal, so it
    // takes the category-side source format.
    let has_category_axis = document
        .descendants()
        .any(|node| node.has_tag_name("catAx") || node.has_tag_name("dateAx"));
    let mut x_axis = None;
    let mut left_axes: Vec<AxisInfo> = Vec::new();
    // (info, is value axis) — only value axes qualify as the secondary scale.
    let mut right_axes: Vec<(AxisInfo, bool)> = Vec::new();
    for axis in document.descendants().filter(|node| {
        ["catAx", "dateAx", "valAx"]
            .iter()
            .any(|name| node.has_tag_name(*name))
    }) {
        let position = direct_child(axis, "axPos").and_then(|node| node.attribute("val"));
        let is_x = match position {
            Some("b") | Some("t") => true,
            Some("l") | Some("r") => false,
            _ => axis.has_tag_name("catAx") || axis.has_tag_name("dateAx"),
        };
        // An axis scaled by another group's series (combo secondary axis)
        // links to that series' cells, not the first series'.
        let own_formats = axis_series(document, axis)
            .filter(|series| Some(*series) != first_series)
            .map(|series| LinkedFormats::from_first_series(Some(series), lookup));
        let formats = own_formats.as_ref().unwrap_or(linked);
        let scaling = direct_child(axis, "scaling");
        let bound = |name: &str| {
            scaling
                .and_then(|node| direct_child(node, name))
                .and_then(|node| node.attribute("val"))
                .and_then(|value| value.parse::<f64>().ok())
        };
        let font_size = |node: Node<'_, '_>| {
            node.descendants()
                .filter(|child| child.has_tag_name("defRPr") || child.has_tag_name("rPr"))
                .find_map(|child| child.attribute("sz"))
                .and_then(|value| value.parse::<f64>().ok())
                .map(|value| value / 100.0)
        };
        let label_props = direct_child(axis, "txPr").and_then(|txpr| {
            txpr.descendants()
                .find(|child| child.has_tag_name("defRPr"))
        });
        let display_units = direct_child(axis, "dispUnits");
        let display_unit = display_units.and_then(|units| {
            direct_child(units, "builtInUnit")
                .and_then(|node| node.attribute("val"))
                .and_then(builtin_display_unit)
                .map(|(divisor, _)| divisor)
                .or_else(|| {
                    direct_child(units, "custUnit")
                        .and_then(|node| node.attribute("val"))
                        .and_then(|value| value.parse::<f64>().ok())
                        .filter(|value| *value > 0.0)
                })
        });
        let info = AxisInfo {
            // Rich text, else the cell-linked strCache <c:v>; a truly empty
            // <c:title> is Excel's auto axis title — the "Axis Title"
            // placeholder.
            title: direct_child(axis, "title").map(|node| {
                let rich = node
                    .descendants()
                    .filter(|child| child.has_tag_name("t"))
                    .filter_map(|child| child.text())
                    .collect::<String>();
                if !rich.is_empty() {
                    return rich;
                }
                let cached = node
                    .descendants()
                    .filter(|child| child.has_tag_name("v"))
                    .filter_map(|child| child.text())
                    .collect::<String>();
                if cached.is_empty() {
                    "Axis Title".into()
                } else {
                    cached
                }
            }),
            min: bound("min"),
            max: bound("max"),
            major_unit: direct_child(axis, "majorUnit")
                .and_then(|node| node.attribute("val"))
                .and_then(|value| value.parse::<f64>().ok()),
            // Source side by axis kind, not position: a horizontal bar's
            // category axis sits on the left and its value axis below.
            num_fmt: formats
                .resolve(
                    direct_child(axis, "numFmt"),
                    axis.has_tag_name("catAx")
                        || axis.has_tag_name("dateAx")
                        || (is_x && !has_category_axis),
                )
                .filter(|code| !code.is_empty() && code != "General"),
            major_gridlines: direct_child(axis, "majorGridlines").is_some(),
            // CT_Boolean: a bare <c:delete/> means true.
            hidden: direct_child(axis, "delete")
                .is_some_and(|node| !matches!(node.attribute("val"), Some("0") | Some("false"))),
            reversed: scaling
                .and_then(|node| direct_child(node, "orientation"))
                .and_then(|node| node.attribute("val"))
                == Some("maxMin"),
            position: position
                .filter(|side| matches!(*side, "l" | "r" | "t" | "b"))
                .map(ToOwned::to_owned),
            label_size: direct_child(axis, "txPr").and_then(font_size),
            label_color: label_props.and_then(|props| drawing_fill_color(props, colors)),
            title_size: direct_child(axis, "title").and_then(font_size),
            title_color: direct_child(axis, "title").and_then(|title| {
                title
                    .descendants()
                    .filter(|child| child.has_tag_name("defRPr") || child.has_tag_name("rPr"))
                    .find_map(|props| drawing_fill_color(props, colors))
            }),
            display_unit,
            // Excel only draws the unit label when c:dispUnitsLbl is present;
            // rich text wins, else the built-in unit's English name.
            display_unit_label: display_units
                .filter(|_| display_unit.is_some())
                .and_then(|units| direct_child(units, "dispUnitsLbl"))
                .map(|label| {
                    let rich = label
                        .descendants()
                        .filter(|child| child.has_tag_name("t"))
                        .filter_map(|child| child.text())
                        .collect::<String>();
                    if !rich.is_empty() {
                        return rich;
                    }
                    direct_child(display_units.unwrap_or(label), "builtInUnit")
                        .and_then(|node| node.attribute("val"))
                        .and_then(builtin_display_unit)
                        .map_or_else(String::new, |(_, name)| name.to_owned())
                })
                .filter(|text| !text.is_empty()),
        };
        if is_x {
            if x_axis.is_none() {
                x_axis = Some(info);
            }
        } else if position == Some("l") {
            left_axes.push(info);
        } else {
            right_axes.push((info, axis.has_tag_name("valAx")));
        }
    }
    // The left axis is the primary scale regardless of document order; the
    // secondary scale is the first remaining right VALUE axis.
    let mut right_values = right_axes
        .into_iter()
        .filter(|(_, is_value)| *is_value)
        .map(|(info, _)| info);
    let (y_axis, secondary_y_axis) = match left_axes.into_iter().next() {
        Some(left) => (Some(left), right_values.next()),
        None => (right_values.next(), right_values.next()),
    };
    (x_axis, y_axis, secondary_y_axis)
}

/// c:builtInUnit divisor and the label Excel shows for it.
pub(crate) fn builtin_display_unit(value: &str) -> Option<(f64, &'static str)> {
    Some(match value {
        "hundreds" => (1e2, "Hundreds"),
        "thousands" => (1e3, "Thousands"),
        "tenThousands" => (1e4, "x 10000"),
        "hundredThousands" => (1e5, "x 100000"),
        "millions" => (1e6, "Millions"),
        "tenMillions" => (1e7, "x 10000000"),
        "hundredMillions" => (1e8, "x 100000000"),
        "billions" => (1e9, "Billions"),
        "trillions" => (1e12, "Trillions"),
        _ => return None,
    })
}

/// Scatter plots carry two valAx (X on the bottom, Y on the left); the left
/// one is the value axis the metadata (gridlines/bounds) should describe.
pub(crate) fn value_axis<'a>(document: &'a Document<'a>) -> Option<Node<'a, 'a>> {
    let axes: Vec<_> = document
        .descendants()
        .filter(|node| node.has_tag_name("valAx"))
        .collect();
    axes.iter()
        .find(|axis| {
            direct_child(**axis, "axPos").and_then(|node| node.attribute("val")) == Some("l")
        })
        .or_else(|| axes.first())
        .copied()
}

pub(crate) fn category_axis_format(
    document: &Document<'_>,
    linked: &LinkedFormats,
) -> Option<String> {
    let axis = document
        .descendants()
        .find(|node| node.has_tag_name("catAx") || node.has_tag_name("dateAx"))?;
    linked.resolve(direct_child(axis, "numFmt"), true)
}

pub(crate) fn value_axis_bounds(document: &Document<'_>) -> Option<ValueAxisBounds> {
    let scaling = direct_child(value_axis(document)?, "scaling")?;
    let bound = |name: &str| {
        direct_child(scaling, name)
            .and_then(|node| node.attribute("val"))
            .and_then(|value| value.parse::<f64>().ok())
    };
    let min = bound("min");
    let max = bound("max");
    (min.is_some() || max.is_some()).then_some(ValueAxisBounds { min, max })
}

pub(crate) fn plot_val_attribute(document: &Document<'_>, plot: &str, name: &str) -> Option<u32> {
    let plot = document
        .descendants()
        .find(|node| node.has_tag_name(plot))?;
    direct_child(plot, name)
        .and_then(|node| node.attribute("val"))
        .and_then(|value| value.parse::<u32>().ok())
}

pub(crate) fn legend_position(document: &Document<'_>) -> String {
    let Some(legend) = document
        .descendants()
        .find(|node| node.has_tag_name("legend"))
    else {
        return "none".into();
    };
    match direct_child(legend, "legendPos").and_then(|node| node.attribute("val")) {
        Some("b") => "bottom",
        Some("t") => "top",
        Some("l") => "left",
        // "r", "tr", or absent all render on the right, the OOXML default.
        _ => "right",
    }
    .into()
}

/// The dLbls node the single-mode metadata reads. Plot-level wins, unless it
/// resolves to no labels while the first series' own dLbls shows some —
/// per-series dLbls override the plot default in Excel, so an all-zero plot
/// element must not hide labels a series switched on.
pub(crate) fn data_labels_node<'a>(document: &'a Document<'a>) -> Option<Node<'a, 'a>> {
    let plot_labels = document
        .descendants()
        .find(|node| flat_plot_name(*node).is_some())
        .and_then(|plot| direct_child(plot, "dLbls"));
    let series_labels = document
        .descendants()
        .find(|node| node.has_tag_name("ser"))
        .and_then(|series| direct_child(series, "dLbls"));
    match (plot_labels, series_labels) {
        (Some(plot), Some(series))
            if data_labels_mode(plot) == "none" && data_labels_mode(series) != "none" =>
        {
            Some(series)
        }
        (Some(plot), _) => Some(plot),
        (None, series) => series,
    }
}

pub(crate) fn data_labels_mode(labels: Node<'_, '_>) -> &'static str {
    let shown = |name: &str| {
        direct_child(labels, name)
            .and_then(|node| node.attribute("val"))
            .is_some_and(|value| value == "1" || value == "true")
    };
    if shown("delete") {
        return "none";
    }
    match (shown("showCatName"), shown("showVal"), shown("showPercent")) {
        (true, true, true) => "category-value-percent",
        (true, _, true) => "category-percent",
        (_, _, true) => "percent",
        (_, true, _) => "value",
        _ => "none",
    }
}

pub(crate) fn data_labels(document: &Document<'_>) -> Option<String> {
    Some(data_labels_mode(data_labels_node(document)?).into())
}

pub(crate) fn data_label_position(document: &Document<'_>) -> Option<String> {
    let position = direct_child(data_labels_node(document)?, "dLblPos")?.attribute("val")?;
    match position {
        "ctr" => Some("center".into()),
        "inEnd" => Some("inside-end".into()),
        "outEnd" => Some("outside-end".into()),
        _ => None,
    }
}

pub(crate) fn data_label_format(document: &Document<'_>, linked: &LinkedFormats) -> Option<String> {
    linked.resolve(direct_child(data_labels_node(document)?, "numFmt"), false)
}

pub(crate) fn axis_titles(document: &Document<'_>) -> Option<AxisTitles> {
    let title_of = |names: &[&str]| -> Option<String> {
        let axis = names
            .iter()
            .find_map(|name| document.descendants().find(|node| node.has_tag_name(*name)))?;
        let text = direct_child(axis, "title")?
            .descendants()
            .filter(|node| node.has_tag_name("t"))
            .filter_map(|node| node.text())
            .collect::<String>();
        (!text.is_empty()).then_some(text)
    };
    let category = title_of(&["catAx", "dateAx"]);
    let value = title_of(&["valAx"]);
    if category.is_none() && value.is_none() {
        return None;
    }
    Some(AxisTitles { category, value })
}

pub(crate) fn plot_grouping(document: &Document<'_>) -> Option<String> {
    let plot = document.descendants().find(|node| {
        matches!(
            flat_plot_name(*node),
            Some("barChart") | Some("areaChart") | Some("lineChart")
        )
    })?;
    direct_child(plot, "grouping")
        .and_then(|node| node.attribute("val"))
        .filter(|value| {
            matches!(
                *value,
                "clustered" | "stacked" | "percentStacked" | "standard"
            )
        })
        .map(ToOwned::to_owned)
}

pub(crate) fn parse_chart_series(
    series: Node<'_, '_>,
    index: usize,
    colors: &ColorContext,
    lookup: &mut SourceFormatLookup<'_>,
) -> ChartSeries {
    // Unnamed series get Excel's global Series1..N numbering. A cell-linked
    // name without a strCache keeps its reference for renderer-side lookup.
    let tx = direct_child(series, "tx");
    let cached_name = tx.and_then(first_cached_value);
    let name_ref = match &cached_name {
        Some(_) => None,
        None => tx.and_then(formula_ref),
    };
    let name = cached_name.unwrap_or_else(|| format!("Series{}", index + 1));
    // Explicit series fill/line color, else the theme accent cycle Excel uses
    // for automatic chart colors — keyed by c:idx, not document position.
    let accent_index = direct_child(series, "idx")
        .and_then(|node| node.attribute("val"))
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(index);
    let color = direct_child(series, "spPr")
        .and_then(|sppr| drawing_fill_color(sppr, colors))
        .or_else(|| theme_accent(colors, accent_index % 6 + 1).map(|base| tint_to_hex(base, 0.0)));
    let category_node = direct_child(series, "cat").or_else(|| direct_child(series, "xVal"));
    let categories = category_node
        .map(|node| {
            cached_points(node)
                .into_iter()
                .map(Option::unwrap_or_default)
                .collect()
        })
        .unwrap_or_default();
    let categories_ref = category_node.and_then(formula_ref);
    // No numCache formatCode (Numbers, openpyxl): the source cells' format is
    // what Excel would show on linked labels and axes.
    let category_format = category_node.and_then(cache_format_code).or_else(|| {
        category_node
            .filter(is_numeric_ref)
            .and(categories_ref.as_deref())
            .and_then(|reference| lookup(reference))
    });
    let value_node = direct_child(series, "val").or_else(|| direct_child(series, "yVal"));
    let values_ref = value_node.and_then(formula_ref);
    let value_points = value_node.map(cached_points).unwrap_or_default();
    let mut values = Vec::with_capacity(value_points.len());
    let mut blanks = Vec::new();
    for (point_index, point) in value_points.iter().enumerate() {
        match point
            .as_deref()
            .and_then(|value| value.trim().parse::<f64>().ok())
            .filter(|value| value.is_finite())
        {
            Some(value) => values.push(value),
            // Blank or non-numeric (#N/A) cache slot: hold the position with
            // a 0 and let the renderer apply c:dispBlanksAs.
            None => {
                values.push(0.0);
                blanks.push(point_index);
            }
        }
    }
    let number_format = value_node.and_then(cache_format_code).or_else(|| {
        values_ref
            .as_deref()
            .and_then(|reference| lookup(reference))
    });
    let trendline = series
        .descendants()
        .find(|node| node.has_tag_name("trendlineType"))
        .and_then(|node| node.attribute("val"))
        .map(ToOwned::to_owned);
    let explosion_pct = direct_child(series, "explosion")
        .and_then(|node| node.attribute("val"))
        .and_then(|value| value.parse::<u32>().ok());
    let point_colors = data_points(series)
        .filter_map(|(index, point)| {
            Some(PointColor {
                index,
                color: direct_child(point, "spPr")
                    .and_then(|sppr| drawing_fill_color(sppr, colors))?,
            })
        })
        .collect::<Vec<_>>();
    let point_explosions = data_points(series)
        .filter_map(|(index, point)| {
            Some(PointExplosion {
                index,
                pct: direct_child(point, "explosion")?
                    .attribute("val")?
                    .parse::<u32>()
                    .ok()?,
            })
        })
        .collect::<Vec<_>>();
    let line = direct_child(series, "spPr")
        .and_then(|sppr| sppr.children().find(|node| node.has_tag_name("ln")));
    // 3-D line/area series are filled ribbons whose outline is noFill by
    // design; folded onto the flat pipeline they must still stroke with
    // their fill color, so the noFill outline is not an empty line here.
    let filled_ribbon = series
        .parent()
        .is_some_and(|plot| plot.has_tag_name("line3DChart") || plot.has_tag_name("area3DChart"));
    let line_color = line.and_then(|ln| {
        if ln.children().any(|node| node.has_tag_name("noFill")) {
            return (!filled_ribbon).then(|| "none".into());
        }
        drawing_fill_color(ln, colors)
    });
    let line_width = line
        .and_then(|ln| ln.attribute("w"))
        .and_then(|value| value.parse::<f64>().ok())
        .map(|emu| emu / 12700.0 * (96.0 / 72.0));
    let smooth = direct_child(series, "smooth")
        .and_then(|node| node.attribute("val"))
        .map(|value| value == "1" || value == "true");
    let marker = direct_child(series, "marker")
        .and_then(|node| direct_child(node, "symbol"))
        .and_then(|node| node.attribute("val"))
        .map(ToOwned::to_owned);
    let plot = series
        .parent()
        .and_then(flat_plot_name)
        .map(ToOwned::to_owned);
    let series_labels = direct_child(series, "dLbls");
    let data_labels = series_labels
        .or_else(|| series.parent().and_then(|plot| direct_child(plot, "dLbls")))
        .map(|labels| data_labels_mode(labels).to_owned());
    let point_labels = series_labels
        .map(point_labels)
        .filter(|labels| !labels.is_empty());
    ChartSeries {
        name,
        name_ref,
        categories,
        values,
        blanks: (!blanks.is_empty()).then_some(blanks),
        number_format,
        category_format,
        color,
        trendline,
        values_ref,
        categories_ref,
        point_colors: (!point_colors.is_empty()).then_some(point_colors),
        explosion_pct,
        point_explosions: (!point_explosions.is_empty()).then_some(point_explosions),
        line_color,
        line_width,
        smooth,
        marker,
        category_groups: category_node.and_then(category_groups),
        plot,
        data_labels,
        point_labels,
    }
}

/// `c:dLbl` entries of a dLbls node. manualLayout x/y default to
/// `factor` mode: an offset from the default label anchor.
pub(crate) fn point_labels(labels: Node<'_, '_>) -> Vec<PointLabel> {
    labels
        .children()
        .filter(|node| node.has_tag_name("dLbl"))
        .filter_map(|label| {
            let index = direct_child(label, "idx")?.attribute("val")?.parse().ok()?;
            let flag = |name: &str| {
                direct_child(label, name)
                    .and_then(|node| node.attribute("val"))
                    .map(|value| value == "1" || value == "true")
            };
            let show_val = if flag("delete") == Some(true) {
                Some(false)
            } else {
                flag("showVal")
            };
            let layout =
                direct_child(label, "layout").and_then(|node| direct_child(node, "manualLayout"));
            let offset = |name: &str| {
                layout
                    .and_then(|node| direct_child(node, name))
                    .and_then(|node| node.attribute("val"))
                    .and_then(|value| value.parse::<f64>().ok())
                    .filter(|value| value.is_finite())
            };
            Some(PointLabel {
                index,
                show_val,
                offset_x: offset("x"),
                offset_y: offset("y"),
            })
        })
        .collect()
}

pub(crate) fn cache_format_code(node: Node<'_, '_>) -> Option<String> {
    node.descendants()
        .find(|child| child.has_tag_name("formatCode"))
        .and_then(|child| child.text())
        .map(ToOwned::to_owned)
}

/// `c:dPt` entries paired with their `c:idx` value.
pub(crate) fn data_points<'a>(series: Node<'a, 'a>) -> impl Iterator<Item = (u32, Node<'a, 'a>)> {
    series
        .children()
        .filter(|node| node.has_tag_name("dPt"))
        .filter_map(|point| {
            let index = direct_child(point, "idx")?
                .attribute("val")?
                .parse::<u32>()
                .ok()?;
            Some((index, point))
        })
}

pub(crate) fn formula_ref(node: Node<'_, '_>) -> Option<String> {
    node.descendants()
        .find(|child| child.has_tag_name("f"))
        .and_then(|child| child.text())
        .map(ToOwned::to_owned)
        .filter(|value| !value.is_empty())
}

pub(crate) fn cached_values(node: Node<'_, '_>) -> Vec<String> {
    // multiLvlStrCache: sweeping every c:pt would flatten L levels into one
    // L×N array. The first c:lvl is the innermost (detail) level per OOXML.
    let scope = node
        .descendants()
        .find(|child| child.has_tag_name("lvl"))
        .unwrap_or(node);
    scope
        .descendants()
        .filter(|child| child.has_tag_name("pt"))
        .filter_map(point_value)
        .collect()
}

pub(crate) fn point_value(point: Node<'_, '_>) -> Option<String> {
    point
        .children()
        .find(|child| child.has_tag_name("v"))
        .and_then(|value| value.text())
        .map(ToOwned::to_owned)
}

/// Caps idx expansion so a corrupt ptCount cannot balloon the arrays.
pub(crate) const MAX_CACHE_SLOTS: usize = 4096;

/// Slot count for idx-aligned expansion: max(ptCount, max idx + 1). None
/// (an idx-less point or an implausible width) falls back to document-order
/// compaction, the legacy behavior.
pub(crate) fn cache_slot_count(cache: Node<'_, '_>, points: &[Node<'_, '_>]) -> Option<usize> {
    let mut max_idx = 0usize;
    for point in points {
        max_idx = max_idx.max(numeric_attribute(*point, "idx")?);
    }
    let pt_count = cache
        .descendants()
        .find(|child| child.has_tag_name("ptCount"))
        .and_then(|count| numeric_attribute(count, "val"))
        .unwrap_or(0);
    let slots = pt_count.max(max_idx + 1);
    (slots <= MAX_CACHE_SLOTS).then_some(slots)
}

/// Cache points expanded onto their `c:pt/@idx` slots (None = blank cell).
/// Sparse category and value caches only stay aligned in idx space —
/// compacting each independently shifts later points left (#prod combo
/// charts whose tail rows are empty).
pub(crate) fn cached_points(node: Node<'_, '_>) -> Vec<Option<String>> {
    let scope = node
        .descendants()
        .find(|child| child.has_tag_name("lvl"))
        .unwrap_or(node);
    let points: Vec<Node<'_, '_>> = scope
        .descendants()
        .filter(|child| child.has_tag_name("pt"))
        .collect();
    if points.is_empty() {
        return Vec::new();
    }
    match cache_slot_count(node, &points) {
        Some(slots) => {
            let mut expanded: Vec<Option<String>> = vec![None; slots];
            for point in &points {
                if let Some(slot) =
                    numeric_attribute(*point, "idx").and_then(|index| expanded.get_mut(index))
                {
                    *slot = point_value(*point);
                }
            }
            expanded
        }
        None => points.iter().map(|point| point_value(*point)).collect(),
    }
}

/// First outer level of a multiLvlStrCache: each pt idx marks a group start
/// in cache-idx space; the span runs until the next pt (or ptCount). Spans
/// are remapped onto the positions `cached_points` emits — idx space when
/// the cache expands, document order otherwise — so sparse caches cannot
/// misalign groups against categories. Ambiguity — an outer pt without an
/// idx, or a group whose members are not contiguous in the emitted order —
/// skips the groups entirely.
pub(crate) fn category_groups(node: Node<'_, '_>) -> Option<Vec<CategoryGroup>> {
    let cache = node
        .descendants()
        .find(|child| child.has_tag_name("multiLvlStrCache"))?;
    let mut levels = cache.children().filter(|child| child.has_tag_name("lvl"));
    let inner = levels.next()?;
    let outer = levels.next()?;
    let inner_points: Vec<Node<'_, '_>> = inner
        .children()
        .filter(|child| child.has_tag_name("pt"))
        .collect();
    if inner_points.is_empty() {
        return None;
    }
    // Mirror cached_points: an expanded cache emits each pt at its idx
    // (blank slots included), a compacted one at its document-order position.
    let slot_count = cache_slot_count(node, &inner_points);
    let inner_idx: Vec<usize> = inner_points
        .iter()
        .enumerate()
        .map(|(position, point)| numeric_attribute(*point, "idx").unwrap_or(position))
        .collect();
    let outer_points: Vec<Node<'_, '_>> = outer
        .children()
        .filter(|child| child.has_tag_name("pt"))
        .collect();
    let mut starts: Vec<(usize, String)> = outer_points
        .iter()
        .filter_map(|point| {
            Some((
                numeric_attribute(*point, "idx")?,
                direct_child(*point, "v")?.text()?.to_owned(),
            ))
        })
        .collect();
    if starts.is_empty() || starts.len() != outer_points.len() {
        return None;
    }
    starts.sort_by_key(|(idx, _)| *idx);
    let point_count = direct_child(cache, "ptCount")
        .and_then(|count| numeric_attribute(count, "val"))
        .unwrap_or_else(|| inner_idx.iter().max().map_or(0, |max| max + 1));
    // A group's members are the emitted positions whose idx falls in
    // [start, next); document order need not be idx order, so require the
    // member positions to be contiguous before expressing them as a span.
    let mut groups = Vec::new();
    for (position, (idx, label)) in starts.iter().enumerate() {
        let next = starts
            .get(position + 1)
            .map_or_else(|| point_count.max(*idx), |(next_idx, _)| *next_idx);
        // Expanded caches carry blank slots too, so the span is the idx
        // range itself.
        if let Some(slots) = slot_count {
            let end = next.min(slots);
            if *idx < end {
                groups.push(CategoryGroup {
                    label: label.clone(),
                    start: *idx,
                    end,
                });
            }
            continue;
        }
        let members: Vec<usize> = inner_idx
            .iter()
            .enumerate()
            .filter(|(_, inner)| (*idx..next).contains(*inner))
            .map(|(emitted, _)| emitted)
            .collect();
        let (Some(&start), Some(&last)) = (members.first(), members.last()) else {
            continue;
        };
        let end = last + 1;
        if end - start != members.len() {
            return None;
        }
        groups.push(CategoryGroup {
            label: label.clone(),
            start,
            end,
        });
    }
    (!groups.is_empty()).then_some(groups)
}

pub(crate) fn first_cached_value(node: Node<'_, '_>) -> Option<String> {
    cached_values(node).into_iter().next().or_else(|| {
        node.descendants()
            .find(|child| child.has_tag_name("v"))
            .and_then(|child| child.text())
            .map(ToOwned::to_owned)
    })
}

pub(crate) fn marker_value(marker: Node<'_, '_>, name: &str) -> Option<usize> {
    marker
        .children()
        .find(|child| child.has_tag_name(name))
        .and_then(|child| child.text())
        .and_then(|value| value.parse::<usize>().ok())
}

pub(crate) fn marker_signed_value(marker: Node<'_, '_>, name: &str) -> Option<i64> {
    marker
        .children()
        .find(|child| child.has_tag_name(name))
        .and_then(|child| child.text())
        .and_then(|value| value.parse::<i64>().ok())
}

pub(crate) fn numeric_attribute(node: Node<'_, '_>, name: &str) -> Option<usize> {
    node.attribute(name)?.parse::<usize>().ok()
}

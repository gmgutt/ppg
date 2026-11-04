package dev.gutt.pptx;

import java.util.List;

/** The content needed to build a {@link PerformanceDeckGenerator} deck. */
public record DeckData(
    String title,
    String subtitle,
    List<String> executiveSummaryPoints,
    String[] years,
    List<PerformanceDeckGenerator.MetricChart> metricCharts) {}

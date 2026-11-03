package dev.gutt.pptx;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.AxisCrossBetween;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.AxisTickMark;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Builds a performance review deck: a title splash, an executive overview, and one chart slide per
 * supplied metric, showing yearly values over the trailing four years.
 */
public final class PerformanceDeckGenerator {

  private static final int SLIDE_WIDTH_PT = 960;
  private static final int SLIDE_HEIGHT_PT = 540;
  private static final Color ACCENT_COLOR = new Color(0x1F, 0x4E, 0x79);
  private static final int YEAR_COUNT = 4;
  private static final Logger log = LoggerFactory.getLogger(PerformanceDeckGenerator.class);

  public record MetricChart(String title, String axisTitle, double[] values) {
    public MetricChart {
      if (values.length != YEAR_COUNT) {
        throw new IllegalArgumentException(
            "expected " + YEAR_COUNT + " yearly values, got " + values.length);
      }
    }
  }

  private final String title;
  private final String subtitle;
  private final List<String> executiveSummaryPoints;
  private final String[] years;
  private final List<MetricChart> metricCharts;

  public PerformanceDeckGenerator(
      String title,
      String subtitle,
      List<String> executiveSummaryPoints,
      String[] years,
      List<MetricChart> metricCharts) {
    if (years.length != YEAR_COUNT) {
      throw new IllegalArgumentException(
          "expected " + YEAR_COUNT + " year labels, got " + years.length);
    }
    if (metricCharts.isEmpty()) {
      throw new IllegalArgumentException("at least one metric chart is required");
    }
    this.title = title;
    this.subtitle = subtitle;
    this.executiveSummaryPoints = List.copyOf(executiveSummaryPoints);
    this.years = years.clone();
    this.metricCharts = List.copyOf(metricCharts);
  }

  public void generate(Path outputFile) throws IOException {
    try (OutputStream out = Files.newOutputStream(outputFile)) {
      generate(out);
    }
  }

  public void generate(OutputStream out) throws IOException {
    try (XMLSlideShow ppt = new XMLSlideShow()) {
      ppt.setPageSize(new Dimension(SLIDE_WIDTH_PT, SLIDE_HEIGHT_PT));
      addTitleSlide(ppt);
      addExecutiveOverviewSlide(ppt);
      for (MetricChart metric : metricCharts) {
        addChartSlide(ppt, metric);
      }
      ppt.write(out);
    }
  }

  private void addTitleSlide(XMLSlideShow ppt) {
    XSLFSlide slide = ppt.createSlide();

    XSLFTextBox titleBox = slide.createTextBox();
    titleBox.setAnchor(new Rectangle2D.Double(60, 200, SLIDE_WIDTH_PT - 120.0, 120));
    XSLFTextParagraph titlePara = titleBox.addNewTextParagraph();
    titlePara.setTextAlign(TextAlign.CENTER);
    XSLFTextRun titleRun = titlePara.addNewTextRun();
    titleRun.setText(title);
    titleRun.setFontSize(40.0);
    titleRun.setBold(true);
    titleRun.setFontColor(ACCENT_COLOR);

    XSLFTextBox subtitleBox = slide.createTextBox();
    subtitleBox.setAnchor(new Rectangle2D.Double(60, 320, SLIDE_WIDTH_PT - 120.0, 60));
    XSLFTextParagraph subtitlePara = subtitleBox.addNewTextParagraph();
    subtitlePara.setTextAlign(TextAlign.CENTER);
    XSLFTextRun subtitleRun = subtitlePara.addNewTextRun();
    subtitleRun.setText(subtitle);
    subtitleRun.setFontSize(20.0);
    subtitleRun.setItalic(true);
    subtitleRun.setFontColor(Color.DARK_GRAY);
  }

  private void addExecutiveOverviewSlide(XMLSlideShow ppt) {
    XSLFSlide slide = ppt.createSlide();
    addSlideHeading(slide, "Executive Overview");

    XSLFTextBox bodyBox = slide.createTextBox();
    bodyBox.setAnchor(
        new Rectangle2D.Double(60, 120, SLIDE_WIDTH_PT - 120.0, SLIDE_HEIGHT_PT - 180.0));
    for (String point : executiveSummaryPoints) {
      XSLFTextParagraph para = bodyBox.addNewTextParagraph();
      para.setBullet(true);
      para.setBulletCharacter("•");
      para.setIndent(18.0);
      para.setLeftMargin(18.0);
      para.setSpaceAfter(12.0);
      XSLFTextRun run = para.addNewTextRun();
      run.setText(point);
      run.setFontSize(18.0);
    }
  }

  private void addChartSlide(XMLSlideShow ppt, MetricChart metric) {
    XSLFSlide slide = ppt.createSlide();
    addSlideHeading(slide, metric.title() + " — " + years[0] + "–" + years[years.length - 1]);

    XSLFChart chart = ppt.createChart();
    double marginPt = 50;
    double topPt = 110;
    Rectangle2D anchorEmu =
        new Rectangle2D.Double(
            Units.toEMU(marginPt),
            Units.toEMU(topPt),
            Units.toEMU(SLIDE_WIDTH_PT - (2 * marginPt)),
            Units.toEMU(SLIDE_HEIGHT_PT - topPt - 40));
    slide.addChart(chart, anchorEmu);

    XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
    bottomAxis.setTitle("Year");
    XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
    leftAxis.setTitle(metric.axisTitle());
    leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
    leftAxis.setMajorTickMark(AxisTickMark.OUT);
    leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

    int n = years.length;
    String categoryRange = chart.formatRange(new CellRangeAddress(1, n, 0, 0));
    String valueRange = chart.formatRange(new CellRangeAddress(1, n, 1, 1));
    XDDFDataSource<String> categoriesData =
        XDDFDataSourcesFactory.fromArray(years, categoryRange, 0);
    XDDFNumericalDataSource<Double> valuesData =
        XDDFDataSourcesFactory.fromArray(boxValues(metric.values()), valueRange, 1);

    XDDFBarChartData bar =
        (XDDFBarChartData) chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
    bar.setBarDirection(BarDirection.COL);
    XDDFBarChartData.Series series =
        (XDDFBarChartData.Series) bar.addSeries(categoriesData, valuesData);
    series.setTitle(metric.title(), chart.setSheetTitle(metric.title(), 1));
    chart.plot(bar);

    XDDFChartLegend legend = chart.getOrAddLegend();
    legend.setPosition(LegendPosition.BOTTOM);
    legend.setOverlay(false);

    chart.setTitleText(metric.title());
    chart.setTitleOverlay(false);
    chart.setAutoTitleDeleted(false);
  }

  private void addSlideHeading(XSLFSlide slide, String text) {
    XSLFTextBox headingBox = slide.createTextBox();
    headingBox.setAnchor(new Rectangle2D.Double(40, 25, SLIDE_WIDTH_PT - 80.0, 60));
    XSLFTextParagraph para = headingBox.addNewTextParagraph();
    XSLFTextRun run = para.addNewTextRun();
    run.setText(text);
    run.setFontSize(28.0);
    run.setBold(true);
    run.setFontColor(ACCENT_COLOR);
  }

  private static Double[] boxValues(double[] values) {
    Double[] boxed = new Double[values.length];
    for (int i = 0; i < values.length; i++) {
      boxed[i] = values[i];
    }
    return boxed;
  }

  /**
   * A single labeled statistic, grouped under a category, with one pre-formatted value per year (in
   * the same order as the years passed to {@link #describeStatistics}).
   */
  public record Statistic(String category, String label, List<String> values) {}

  /**
   * Renders a data set of statistics as a prompt-style text block, grouped by category, framed
   * around describing the most recent year with earlier years as supporting context. This is
   * intended to be handed to an LLM to produce a natural-language narrative of the data.
   */
  public static String describeStatistics(List<String> years, List<Statistic> statistics) {
    if (years.isEmpty()) {
      throw new IllegalArgumentException("at least one year is required");
    }
    if (statistics.isEmpty()) {
      throw new IllegalArgumentException("at least one statistic is required");
    }

    Map<String, List<Statistic>> byCategory = new LinkedHashMap<>();
    for (Statistic statistic : statistics) {
      byCategory.computeIfAbsent(statistic.category(), c -> new ArrayList<>()).add(statistic);
    }

    String yearList = String.join(", ", years);
    String mostRecentYear = years.getLast();

    StringBuilder text = new StringBuilder();
    text.append("Below is a data set with ")
        .append(statistics.size())
        .append(" metrics, each with one value per year\n")
        .append("(")
        .append(yearList)
        .append(", in that order), grouped by the category in\n")
        .append("their labels. Describe the data using only what is in this message.\n")
        .append("The focus is the most recent year (")
        .append(mostRecentYear)
        .append("); earlier years are used only\n")
        .append("as supporting context for ")
        .append(mostRecentYear)
        .append(".\n");

    for (Map.Entry<String, List<Statistic>> entry : byCategory.entrySet()) {
      text.append("\n").append(entry.getKey()).append("\n");
      for (Statistic statistic : entry.getValue()) {
        text.append("- ")
            .append(statistic.label())
            .append(": ")
            .append(String.join(", ", statistic.values()))
            .append("\n");
      }
    }

    return text
        + """
        Below is a data set with 23 metrics, each with one value per year
        (2023, 2024, 2025, 2026, in that order), grouped by the category in
        their labels. Describe the data using only what is in this message.
        The focus is the most recent year (2026); earlier years are used only
        as supporting context for 2026.

        Output format (must fit on one presentation slide, 150 words maximum
        in total):

        - One headline sentence (30 words or fewer) about the 2026 values
          across all categories.
        - Exactly three bullets, one per category (category name first), each
          40 words or fewer, naming the metrics whose 2026 values stand out
          most against earlier years in that category, with the value and
          comparison for each.
        - One closing bullet (15 words or fewer) with the single most notable
          comparison of 2026 against earlier years, across any metric.

        Rules:
        - Evaluate all 23 metrics before choosing what to report, but report
          only the most notable ones. Notable means the largest change from
          2025 to 2026, the highest or lowest 2026 value relative to earlier
          years, or a 2026 value that reverses the earlier direction. Judge
          this by the numbers in this data set only.
        - Every statement must be about 2026. Earlier years may appear only as
          a comparison point for 2026, never as the subject of a statement.
        - Describe direction only as "increased" or "decreased". Do not say
          improved, worsened, better, or worse, since the data does not say
          which direction is favorable for any metric.
        - Every comparison must be between values in this data set. Do not use
          your own knowledge of what is typical, average, good, bad,
          impressive, or unusual for any metric, sport, player, or league.
        - Treat the labels as opaque names. Do not infer what the metrics
          measure, what position is played, or what caused any change.
        - Do not predict future values or assume anything about sample size,
          playing time, or data quality.
        - Output only the headline and four bullets. No introduction,
          explanation, extra sections, or concluding remarks.
        - Report numbers exactly as given. Compute percentage change only where
          the starting value is nonzero.
        - If you cannot fit something within the limits, omit it rather than
          exceeding them.
        """;
  }

  /**
   * Sample data sourced from Pete Crow-Armstrong's Standard Batting, Advanced Batting, and Standard
   * Fielding tables (Chicago Cubs, 2023-2026):
   * <a href="https://www.baseball-reference.com/players/c/crowape01.shtml">...</a>
   */
  public static void main(String[] args) throws IOException {
    String[] years = {"2023", "2024", "2025", "2026"};

    List<MetricChart> metrics =
        List.of(
            // --- original three ---
            new MetricChart("WAR (Standard Batting)", "WAR", new double[] {-0.1, 2.3, 5.9, 9.8}),
            new MetricChart(
                "Isolated Power – ISO (Advanced Batting)",
                "ISO",
                new double[] {0.000, 0.148, 0.234, 0.291}),
            new MetricChart(
                "Defensive Runs Saved (Standard Fielding)", "DRS", new double[] {1, 11, 15, 16}),
            // --- 20 more, Standard Batting ---
            new MetricChart(
                "On-Base Plus Slugging – OPS (Standard Batting)",
                "OPS",
                new double[] {0.176, 0.670, 0.768, 0.943}),
            new MetricChart("OPS+ (Standard Batting)", "OPS+", new double[] {-45, 88, 114, 156}),
            new MetricChart(
                "Batting Average (Standard Batting)",
                "BA",
                new double[] {0.000, 0.237, 0.247, 0.280}),
            new MetricChart(
                "On-Base Percentage (Standard Batting)",
                "OBP",
                new double[] {0.176, 0.286, 0.287, 0.372}),
            new MetricChart(
                "Slugging Percentage (Standard Batting)",
                "SLG",
                new double[] {0.000, 0.384, 0.481, 0.570}),
            new MetricChart("Home Runs (Standard Batting)", "HR", new double[] {0, 10, 31, 45}),
            new MetricChart(
                "Runs Batted In (Standard Batting)", "RBI", new double[] {0, 47, 95, 107}),
            new MetricChart("Stolen Bases (Standard Batting)", "SB", new double[] {1, 27, 35, 41}),
            // --- 20 more, Advanced Batting ---
            new MetricChart(
                "Walk Rate – BB% (Advanced Batting)", "BB%", new double[] {15.8, 5.1, 4.5, 11.4}),
            new MetricChart(
                "Strikeout Rate – SO% (Advanced Batting)",
                "SO%", new double[] {36.8, 23.9, 24.0, 25.8}),
            new MetricChart(
                "Hard-Hit Rate – HardH% (Advanced Batting)",
                "HardH%", new double[] {22.2, 36.8, 41.9, 48.8}),
            new MetricChart(
                "Batting Average on Balls in Play – BAbip (Advanced Batting)",
                "BAbip",
                new double[] {0.000, 0.292, 0.276, 0.327}),
            new MetricChart(
                "Win Probability Added – WPA (Advanced Batting)",
                "WPA",
                new double[] {-0.4, -0.5, 1.1, 5.2}),
            new MetricChart(
                "Stolen Base Success Rate – SB% (Advanced Batting)",
                "SB%", new double[] {50.0, 90.0, 81.4, 85.4}),
            new MetricChart(
                "Average Exit Velocity – EV (Advanced Batting)",
                "EV (mph)",
                new double[] {71.4, 84.6, 88.0, 89.8}),
            // --- 20 more, Standard Fielding (CF) ---
            new MetricChart(
                "Fielding Percentage (Standard Fielding)",
                "Fld%",
                new double[] {1.000, 0.991, 0.995, 0.985}),
            new MetricChart(
                "Range Factor per 9 Innings (Standard Fielding)",
                "RF/9",
                new double[] {2.06, 2.93, 2.87, 2.61}),
            new MetricChart(
                "Total Zone Runs – Rtot (Standard Fielding)", "Rtot", new double[] {2, 3, 5, 5}),
            new MetricChart("Assists (Standard Fielding)", "A", new double[] {0, 8, 5, 5}),
            new MetricChart("Errors (Standard Fielding)", "E", new double[] {0, 3, 2, 6}));

    List<String> executiveSummary =
        List.of(
            "WAR rose from -0.1 in a 13-game 2023 debut to a league-leading 9.8 in 2026, "
                + "reflecting four straight years of growth as a complete player.",
            "Isolated Power (ISO) nearly doubled from .148 in his first full season (2024) to "
                + ".291 in 2026, as his home run total climbed from 10 to 45.",
            "Defensive Runs Saved in center field climbed from 1 to 16 over the same span, "
                + "underpinning elite two-way value.",
            "2026 was a breakout year across the board: a .943 OPS (156 OPS+), 45 HR, 41 SB, "
                + "and an NL All-Star selection.");

    PerformanceDeckGenerator generator =
        new PerformanceDeckGenerator(
            "Pete Crow-Armstrong — Performance Review",
            "Standard Batting, Advanced Batting & Standard Fielding — 2023–2026 (Chicago Cubs)",
            executiveSummary,
            years,
            metrics);

    Path output = Path.of("target", "performance-review.pptx");
    generator.generate(output);
    log.info("Wrote {}", output.toAbsolutePath());
  }
}

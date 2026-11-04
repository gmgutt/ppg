package dev.gutt.pptx;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
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

  public record MetricChart(String title, String axisTitle, List<Double> values) {
    public MetricChart {
      if (values.size() != YEAR_COUNT) {
        throw new IllegalArgumentException(
            "expected " + YEAR_COUNT + " yearly values, got " + values.size());
      }
      values = List.copyOf(values);
    }
  }

  private final String title;
  private final String subtitle;
  private final List<String> executiveSummaryPoints;
  private final List<String> years;
  private final List<MetricChart> metricCharts;

  public PerformanceDeckGenerator(
      String title,
      String subtitle,
      List<String> executiveSummaryPoints,
      List<String> years,
      List<MetricChart> metricCharts) {
    if (years.size() != YEAR_COUNT) {
      throw new IllegalArgumentException(
          "expected " + YEAR_COUNT + " year labels, got " + years.size());
    }
    if (metricCharts.isEmpty()) {
      throw new IllegalArgumentException("at least one metric chart is required");
    }
    this.title = title;
    this.subtitle = subtitle;
    this.executiveSummaryPoints = List.copyOf(executiveSummaryPoints);
    this.years = List.copyOf(years);
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
    addSlideHeading(slide, metric.title() + " — " + years.getFirst() + "–" + years.getLast());

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

    int n = years.size();
    String categoryRange = chart.formatRange(new CellRangeAddress(1, n, 0, 0));
    String valueRange = chart.formatRange(new CellRangeAddress(1, n, 1, 1));
    XDDFDataSource<String> categoriesData =
        XDDFDataSourcesFactory.fromArray(years.toArray(String[]::new), categoryRange, 0);
    XDDFNumericalDataSource<Double> valuesData =
        XDDFDataSourcesFactory.fromArray(metric.values().toArray(Double[]::new), valueRange, 1);

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

  public static void main(String[] args) throws IOException {
    DeckDataService deckDataService = new DeckDataService();
    DeckData data =
        deckDataService
            .findById(DeckDataService.SAMPLE_DECK_ID)
            .orElseThrow(() -> new IllegalStateException("sample deck not found"));

    PerformanceDeckGenerator generator =
        new PerformanceDeckGenerator(
            data.title(),
            data.subtitle(),
            data.executiveSummaryPoints(),
            data.years(),
            data.metricCharts());

    Path output = Path.of("target", "performance-review.pptx");
    generator.generate(output);
    log.info("Wrote {}", output.toAbsolutePath());
  }
}

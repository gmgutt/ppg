package dev.gutt.pptx;

import static io.restassured.RestAssured.given;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.containsString;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.matchesPattern;
import static org.hamcrest.Matchers.stringContainsInOrder;
import static org.junit.jupiter.api.Assertions.assertAll;

import io.quarkus.test.junit.QuarkusTest;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Pattern;
import java.util.stream.DoubleStream;
import java.util.stream.IntStream;
import java.util.stream.Stream;
import org.apache.poi.sl.usermodel.PaintStyle.SolidPaint;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.junit.jupiter.api.Test;

@QuarkusTest
class DeckResourceTest {

  private static final String PPTX_MEDIA_TYPE =
      "application/vnd.openxmlformats-officedocument.presentationml.presentation";

  private final DeckData expected =
      new DeckDataService().findById(DeckDataService.SAMPLE_DECK_ID).orElseThrow();

  @Test
  void getDeckReturnsPptxWithOneSlidePerSectionForKnownId() throws IOException {

    byte[] body =
        given()
            .when()
            .get("/decks/{id}", DeckDataService.SAMPLE_DECK_ID)
            .then()
            .statusCode(200)
            .contentType(PPTX_MEDIA_TYPE)
            .header("Content-Disposition", containsString(DeckDataService.SAMPLE_DECK_ID))
            .extract()
            .asByteArray();

    XMLSlideShow ppt = new XMLSlideShow(new ByteArrayInputStream(body));

    assertThat(ppt.getPageSize(), is(new Dimension(960, 540)));
    // title slide + executive overview + one slide per metric
    assertThat(ppt.getSlides().size(), is(2 + expected.metricCharts().size()));
    assertThat(ppt.getCharts().size(), is(expected.metricCharts().size()));
  }

  @Test
  void getDeckTitleSlideShowsTitleAndSubtitle() throws IOException {

    byte[] body =
        given()
            .when()
            .get("/decks/{id}", DeckDataService.SAMPLE_DECK_ID)
            .then()
            .statusCode(200)
            .contentType(PPTX_MEDIA_TYPE)
            .header("Content-Disposition", containsString(DeckDataService.SAMPLE_DECK_ID))
            .extract()
            .asByteArray();

    XMLSlideShow ppt = new XMLSlideShow(new ByteArrayInputStream(body));

    XSLFSlide slide = ppt.getSlides().getFirst();

    assertThat(slide.getShapes().size(), is(2));

    XSLFTextBox title = (XSLFTextBox) slide.getShapes().getFirst();
    assertThat(title.getAnchor(), is(new Rectangle2D.Double(60, 200, 840, 120)));

    XSLFTextParagraph titlePara = lastParagraph(title);
    assertThat(titlePara.getTextAlign(), is(TextAlign.CENTER));

    XSLFTextRun titleRun = titlePara.getTextRuns().getFirst();
    assertThat(titleRun.getRawText(), is(expected.title()));
    assertThat(titleRun.getFontSize(), is(40.0));
    assertThat(titleRun.isBold(), is(true));
    assertThat(solidColor(titleRun), is(new Color(0x1F, 0x4E, 0x79)));

    XSLFTextBox subtitle = (XSLFTextBox) slide.getShapes().get(1);
    assertThat(subtitle.getAnchor(), is(new Rectangle2D.Double(60, 320, 840, 60)));

    XSLFTextParagraph subtitlePara = lastParagraph(subtitle);
    assertThat(subtitlePara.getTextAlign(), is(TextAlign.CENTER));

    XSLFTextRun subtitleRun = subtitlePara.getTextRuns().getFirst();
    assertThat(subtitleRun.getRawText(), is(expected.subtitle()));
    assertThat(subtitleRun.getFontSize(), is(20.0));
    assertThat(subtitleRun.isItalic(), is(true));
    assertThat(solidColor(subtitleRun), is(Color.DARK_GRAY));
  }

  @Test
  void getDeckExecutiveOverviewSlideListsSummaryPoints() throws IOException {

    byte[] body =
        given()
            .when()
            .get("/decks/{id}", DeckDataService.SAMPLE_DECK_ID)
            .then()
            .statusCode(200)
            .contentType(PPTX_MEDIA_TYPE)
            .header("Content-Disposition", containsString(DeckDataService.SAMPLE_DECK_ID))
            .extract()
            .asByteArray();

    XMLSlideShow ppt = new XMLSlideShow(new ByteArrayInputStream(body));

    XSLFSlide slide = ppt.getSlides().get(1);

    assertHeading(slide, "Executive Overview");

    XSLFTextBox bulletBox = (XSLFTextBox) slide.getShapes().get(1);
    assertThat(bulletBox.getAnchor(), is(new Rectangle2D.Double(60, 120, 840, 360)));

    List<String> points = expected.executiveSummaryPoints();

    // addNewTextParagraph appends after the default empty paragraph
    List<XSLFTextParagraph> paragraphs = bulletBox.getTextParagraphs();
    for (int i = 0; i < points.size(); i++) {
      XSLFTextParagraph para = paragraphs.get(paragraphs.size() - points.size() + i);
      assertThat(para.isBullet(), is(true));
      assertThat(para.getBulletCharacter(), is("•"));
      assertThat(para.getIndent(), is(18.0));
      assertThat(para.getLeftMargin(), is(18.0));
      assertThat(para.getSpaceAfter(), is(12.0));

      XSLFTextRun run = para.getTextRuns().getFirst();
      assertThat(run.getRawText(), is(points.get(i)));
      assertThat(run.getFontSize(), is(18.0));
    }
  }

  @Test
  void getDeckHasOneChartSlidePerMetric() throws IOException {

    byte[] body =
        given()
            .when()
            .get("/decks/{id}", DeckDataService.SAMPLE_DECK_ID)
            .then()
            .statusCode(200)
            .contentType(PPTX_MEDIA_TYPE)
            .header("Content-Disposition", containsString(DeckDataService.SAMPLE_DECK_ID))
            .extract()
            .asByteArray();

    XMLSlideShow ppt = new XMLSlideShow(new ByteArrayInputStream(body));

    assertAll(
        IntStream.range(0, expected.metricCharts().size())
            .mapToObj(i ->
                () -> {
                  XSLFSlide slide = ppt.getSlides().get(i + 2);
                  PerformanceDeckGenerator.MetricChart metric = expected.metricCharts().get(i);

                  String years =
                      expected.years()[0] + "–" + expected.years()[expected.years().length - 1];
                  assertHeading(slide, metric.title() + " — " + years);

                  XSLFGraphicFrame frame = (XSLFGraphicFrame) slide.getShapes().get(1);
                  assertThat(frame.getAnchor(), is(new Rectangle2D.Double(50, 110, 860, 390)));

                  XSLFChart chart = frame.getChart();
                  String xml = chart.getCTChartSpace().xmlText();
                  // chart title, shown (not overlaid, not auto-deleted)
                  assertThat(chart.getTitleShape().getText(), is(metric.title()));
                  assertThat(
                      xml,
                      matchesPattern(
                          "(?s).*<c:chart><c:title><c:tx>(?:(?!</c:title>).)*"
                              + "<c:overlay val=\"false\"/></c:title><c:autoTitleDeleted val=\"false\"/>.*"));
                  // one column series: title and one value per year
                  assertThat(xml, containsString("<c:barDir val=\"col\"/>"));
                  assertThat(
                      chart
                          .getCTChart()
                          .getPlotArea()
                          .getBarChartArray(0)
                          .getSerArray(0)
                          .getTx()
                          .getStrRef()
                          .getStrCache()
                          .getPtArray(0)
                          .getV(),
                      is(metric.title()));
                  // categories (years) come first in the XML, followed by the series values
                  assertThat(
                      xml,
                      stringContainsInOrder(
                          Stream.concat(
                                  Arrays.stream(expected.years()),
                                  DoubleStream.of(metric.values()).mapToObj(String::valueOf))
                              .map(v -> "<c:v>" + v + "</c:v>")
                              .toList()));
                  // category axis
                  assertThat(
                      xml,
                      matchesPattern(
                          "(?s).*<c:catAx>(?:(?!</c:catAx>).)*<a:t>Year</a:t>.*</c:catAx>.*"));
                  // value axis
                  assertThat(
                      xml,
                      matchesPattern(
                          "(?s).*<c:valAx>(?:(?!</c:valAx>).)*<a:t>"
                              + Pattern.quote(metric.axisTitle())
                              + "</a:t>(?:(?!</c:valAx>).)*<c:majorTickMark val=\"out\"/>"
                              + "(?:(?!</c:valAx>).)*<c:crosses val=\"autoZero\"/>"
                              + "<c:crossBetween val=\"between\"/></c:valAx>.*"));
                  // legend at the bottom, not overlaid
                  assertThat(
                      xml,
                      containsString(
                          "<c:legend><c:legendPos val=\"b\"/><c:overlay val=\"false\"/></c:legend>"));
                }));
  }

  @Test
  void getDeckReturnsNotFoundForUnknownId() {
    given().when().get("/decks/{id}", "does-not-exist").then().statusCode(404);
  }

  private static void assertHeading(XSLFSlide slide, String text) {

    XSLFTextBox heading = (XSLFTextBox) slide.getShapes().getFirst();
    assertThat(heading.getAnchor(), is(new Rectangle2D.Double(40, 25, 880, 60)));

    XSLFTextRun run = lastParagraph(heading).getTextRuns().getFirst();
    assertThat(run.getRawText(), is(text));
    assertThat(run.getFontSize(), is(28.0));
    assertThat(run.isBold(), is(true));
    assertThat(solidColor(run), is(new Color(0x1F, 0x4E, 0x79)));
  }

  /** addNewTextParagraph appends after the text box's default empty paragraph. */
  private static XSLFTextParagraph lastParagraph(XSLFTextBox box) {
    return box.getTextParagraphs().getLast();
  }

  private static Color solidColor(XSLFTextRun run) {
    return ((SolidPaint) run.getFontColor()).getSolidColor().getColor();
  }
}

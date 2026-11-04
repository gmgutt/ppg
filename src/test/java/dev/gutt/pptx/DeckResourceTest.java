package dev.gutt.pptx;

import static io.restassured.RestAssured.given;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.containsString;
import static org.hamcrest.Matchers.empty;
import static org.hamcrest.Matchers.not;

import io.quarkus.test.junit.QuarkusTest;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.junit.jupiter.api.Test;

@QuarkusTest
class DeckResourceTest {

  private static final String PPTX_MEDIA_TYPE =
      "application/vnd.openxmlformats-officedocument.presentationml.presentation";

  @Test
  void getDeckReturnsPptxForKnownId() throws IOException {
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

    try (XMLSlideShow ppt = new XMLSlideShow(new ByteArrayInputStream(body))) {
      assertThat(ppt.getSlides(), not(empty()));
    }
  }

  @Test
  void getDeckReturnsNotFoundForUnknownId() {
    given().when().get("/decks/{id}", "does-not-exist").then().statusCode(404);
  }
}

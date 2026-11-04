package dev.gutt.pptx;

import jakarta.ws.rs.GET;
import jakarta.ws.rs.NotFoundException;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.PathParam;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.Response;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

@Path("/decks")
public class DeckResource {

  private static final String PPTX_MEDIA_TYPE =
      "application/vnd.openxmlformats-officedocument.presentationml.presentation";

  DeckDataService deckDataService;

  public DeckResource(DeckDataService deckDataService) {
    this.deckDataService = deckDataService;
  }

  @GET
  @Path("/{id}")
  @Produces(PPTX_MEDIA_TYPE)
  public Response getDeck(@PathParam("id") String id) throws IOException {
    DeckData data = deckDataService.findById(id).orElseThrow(NotFoundException::new);

    PerformanceDeckGenerator generator =
        new PerformanceDeckGenerator(
            data.title(),
            data.subtitle(),
            data.executiveSummaryPoints(),
            data.years(),
            data.metricCharts());

    ByteArrayOutputStream out = new ByteArrayOutputStream();
    generator.generate(out);

    return Response.ok(out.toByteArray())
        .header("Content-Disposition", "attachment; filename=\"" + id + ".pptx\"")
        .build();
  }
}

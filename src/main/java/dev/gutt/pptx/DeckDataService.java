package dev.gutt.pptx;

import dev.gutt.pptx.PerformanceDeckGenerator.MetricChart;
import jakarta.enterprise.context.ApplicationScoped;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/** Looks up the data behind a performance deck by id. */
@ApplicationScoped
public class DeckDataService {

  /**
   * Sourced from Pete Crow-Armstrong's Standard Batting, Advanced Batting, and Standard Fielding
   * tables (Chicago Cubs, 2023-2026): <a
   * href="https://www.baseball-reference.com/players/c/crowape01.shtml">...</a>
   */
  public static final String SAMPLE_DECK_ID = "pete-crow-armstrong";

  private final Map<String, DeckData> decksById = Map.of(SAMPLE_DECK_ID, sampleDeck());

  public Optional<DeckData> findById(String id) {
    return Optional.ofNullable(decksById.get(id));
  }

  private static DeckData sampleDeck() {
    List<String> years = List.of("2023", "2024", "2025", "2026");

    List<MetricChart> metrics =
        List.of(
            // --- original three ---
            new MetricChart("WAR (Standard Batting)", "WAR", List.of(-0.1, 2.3, 5.9, 9.8)),
            new MetricChart(
                "Isolated Power – ISO (Advanced Batting)",
                "ISO",
                List.of(0.000, 0.148, 0.234, 0.291)),
            new MetricChart(
                "Defensive Runs Saved (Standard Fielding)", "DRS", List.of(1.0, 11.0, 15.0, 16.0)),
            // --- 20 more, Standard Batting ---
            new MetricChart(
                "On-Base Plus Slugging – OPS (Standard Batting)",
                "OPS",
                List.of(0.176, 0.670, 0.768, 0.943)),
            new MetricChart("OPS+ (Standard Batting)", "OPS+", List.of(-45.0, 88.0, 114.0, 156.0)),
            new MetricChart(
                "Batting Average (Standard Batting)", "BA", List.of(0.000, 0.237, 0.247, 0.280)),
            new MetricChart(
                "On-Base Percentage (Standard Batting)",
                "OBP",
                List.of(0.176, 0.286, 0.287, 0.372)),
            new MetricChart(
                "Slugging Percentage (Standard Batting)",
                "SLG",
                List.of(0.000, 0.384, 0.481, 0.570)),
            new MetricChart("Home Runs (Standard Batting)", "HR", List.of(0.0, 10.0, 31.0, 45.0)),
            new MetricChart(
                "Runs Batted In (Standard Batting)", "RBI", List.of(0.0, 47.0, 95.0, 107.0)),
            new MetricChart(
                "Stolen Bases (Standard Batting)", "SB", List.of(1.0, 27.0, 35.0, 41.0)),
            // --- 20 more, Advanced Batting ---
            new MetricChart(
                "Walk Rate – BB% (Advanced Batting)", "BB%", List.of(15.8, 5.1, 4.5, 11.4)),
            new MetricChart(
                "Strikeout Rate – SO% (Advanced Batting)", "SO%", List.of(36.8, 23.9, 24.0, 25.8)),
            new MetricChart(
                "Hard-Hit Rate – HardH% (Advanced Batting)",
                "HardH%", List.of(22.2, 36.8, 41.9, 48.8)),
            new MetricChart(
                "Batting Average on Balls in Play – BAbip (Advanced Batting)",
                "BAbip",
                List.of(0.000, 0.292, 0.276, 0.327)),
            new MetricChart(
                "Win Probability Added – WPA (Advanced Batting)",
                "WPA",
                List.of(-0.4, -0.5, 1.1, 5.2)),
            new MetricChart(
                "Stolen Base Success Rate – SB% (Advanced Batting)",
                "SB%", List.of(50.0, 90.0, 81.4, 85.4)),
            new MetricChart(
                "Average Exit Velocity – EV (Advanced Batting)",
                "EV (mph)",
                List.of(71.4, 84.6, 88.0, 89.8)),
            // --- 20 more, Standard Fielding (CF) ---
            new MetricChart(
                "Fielding Percentage (Standard Fielding)",
                "Fld%",
                List.of(1.000, 0.991, 0.995, 0.985)),
            new MetricChart(
                "Range Factor per 9 Innings (Standard Fielding)",
                "RF/9",
                List.of(2.06, 2.93, 2.87, 2.61)),
            new MetricChart(
                "Total Zone Runs – Rtot (Standard Fielding)", "Rtot", List.of(2.0, 3.0, 5.0, 5.0)),
            new MetricChart("Assists (Standard Fielding)", "A", List.of(0.0, 8.0, 5.0, 5.0)),
            new MetricChart("Errors (Standard Fielding)", "E", List.of(0.0, 3.0, 2.0, 6.0)));

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

    return new DeckData(
        "Pete Crow-Armstrong — Performance Review",
        "Standard Batting, Advanced Batting & Standard Fielding — 2023–2026 (Chicago Cubs)",
        executiveSummary,
        years,
        metrics);
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
}

package dev.gutt.pptx;

import dev.gutt.pptx.PerformanceDeckGenerator.MetricChart;
import jakarta.enterprise.context.ApplicationScoped;
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

    return new DeckData(
        "Pete Crow-Armstrong — Performance Review",
        "Standard Batting, Advanced Batting & Standard Fielding — 2023–2026 (Chicago Cubs)",
        executiveSummary,
        years,
        metrics);
  }
}

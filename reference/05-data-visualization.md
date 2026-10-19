# 05 — Data Visualization Rules

Distilled from PlacePlate's data guide + academic guide. Core belief: **raw
numbers don't persuade — visual stories do.** Every data slide must make its
finding impossible to miss.

## Step 0 — ask what comparison you're making

| Question the data answers | Chart |
|---|---|
| Comparing categories | Bar/column chart (horizontal bars when labels are long) |
| Change over time | Line chart |
| Part-to-whole | Donut, **≤ 4 segments**, labels on the chart (no legend) |
| Distribution / correlation | Scatter plot |
| Exact lookup values, 4+ variables | Table (designed, see below) |
| One number carries the message | Hero stat (R9), no chart at all |

**Never:** 3D charts (distort values, look 2003); pie charts beyond 3 segments
(eyes can't compare angles — a bar chart always wins at 4+).

## Declutter protocol (apply to every chart)

PowerPoint's default chart is noise. Strip it:

1. **Kill the legend** for single-series charts; for multi-series, label lines/
   bars directly — no back-and-forth eye travel.
2. **Gridlines** very light gray or gone (gone if you have data labels).
3. **Direct data labels** on bars/points, in the deck's body font — exact values
   readable without the axis.
4. **No chart border/background** — chart area melts into the slide background;
   data floats on the slide.
5. **Axis starts at zero** for bars. Truncated axes = lying (log scale excepted,
   and say so).
6. Title inside the chart object: **off**. The *slide headline* is the chart title.

## Color = the argument ("strategic gray")

- Color every bar/line **gray** except the one that carries your point → accent.
  Gray says "context"; accent says "finding". The audience gets it before you speak.
- **One accent color per chart.** If everything is colorful, nothing stands out.
- Red = negative/warning only. Green = positive/growth only. (Cultural semantics
  are load-bearing — don't fight them.)
- Brand accent goes on the data point you want remembered.

## Headlines state findings

- ✕ "Revenue by Quarter" → ✓ "Q3 revenue hit a 3-year high"
- ✕ "Survey Results" → ✓ "78% of users churn over slow onboarding"
- ✕ "Market Overview" → ✓ "Our segment grows 2× faster than the market"

The story must land before a single data point is read. Pair the chart with a
right-rail **insight card** (R8): 1–2 sentences of "so what", so the takeaway
survives even if nobody studies the axes.

## Hero stat construction

1. Number **80–120 pt**, accent color ("127%", "$4.2M").
2. One-line label ~24 pt ("Year-over-year revenue growth").
3. Context line ~18 pt gray ("vs 34% industry average") — one comparison that
   makes the number mean something.

Place at **section breaks**: after detail-dense slides it resets attention and
stamps the "so what". Understandable in 3 seconds or redesign it.

## Tables (when you must)

Appropriate only for exact multi-value lookup or 4+ variable comparisons.

- **Max 5 columns × 7 rows** on a slide. (Academic: show 3–4 columns max —
  never paste a raw 10-column spreadsheet; rebuild it.)
- Alternating row shading (white / very light gray), **no cell borders**.
- Header row in accent with white text.
- **Highlight the key cell** (accent bold) so the conclusion is pre-computed.
- If the table exists to show a trend, you're using the wrong tool — chart it.

## Data-slide failure modes

| Mistake | Why it fails | Fix |
|---|---|---|
| Y-axis not at zero | Small deltas look dramatic (reads as dishonest) | Start at 0 |
| 8 lines on one chart | Unreadable spaghetti | Small multiples or split slides |
| Pie with 6 slices | Angle comparison is guesswork | Bar chart |
| Topic-label title | Audience must derive the point themselves | Assertion headline |
| Raw Excel paste | Illegible beyond row 1 (front) | Rebuild: 3–4 cols, key rows only |
| Data with no source (academic/pitch) | Credibility dies on first question | Cite source + date on-slide, small but visible |
| Unlabeled axes/units | Numbers without meaning | Label axes, include units |

## Morph + data (their signature data move)

Duplicate the chart slide → change values → Morph animates growth/shrinkage.
Sequencing patterns: baseline year → current year (bars rise); hero stat → full
chart (number becomes its highlighted bar); donut segment → its own detail slide.
Keep the chart object *the same object* across duplicates (same position/size)
so matching is exact.

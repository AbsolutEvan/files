# Visualization source

The Sankey / alluvial diagram lives here (self-contained HTML + D3).

Planned pieces:

- **Alluvial layout** — nodes are (year, grade) buckets; each school is a
  ribbon threading through one node per year.
- **Ribbon width** = that year's enrollment.
- **Segment color** = change in the selected subject's passing rate across
  that year-to-year segment (green = improved, red = declined, gray = flat).
- **ELA / Math toggle** — swaps which subject drives segment colors.
- **School lookup** — search box highlights one school's full path.

Reads `../data/processed/schools.json`.

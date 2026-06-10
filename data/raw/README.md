# Raw data — drop your files here

Put the original `.xlsx` files in this folder, exactly as you have them.
Don't bother cleaning them up first — that's what the processing step is for.

A few things that help:

- **One file per year is fine, or one file with a year column — either works.**
  Just let me know which.
- **Keep the original filenames** if they encode the year (e.g. `grades_2019.xlsx`).
- If a workbook has multiple sheets, tell me which sheet holds the data
  (or I'll inspect and ask).

### What I'll pull out of each row

| field        | notes                                              |
|--------------|----------------------------------------------------|
| school_id    | stable unique ID (preferred over name)             |
| school_name  | for the lookup/search box                          |
| year         | school year (ending/spring year by convention)     |
| grade        | A / B / C / D / F                                  |
| ela_pass_rate| ELA standardized-test passing rate                 |
| math_pass_rate| Math standardized-test passing rate               |
| enrollment   | enrollment **for that year** (drives ribbon width) |

The column names in your spreadsheets don't need to match these — I'll map
them during cleaning. Just make sure the underlying numbers are in here
somewhere.

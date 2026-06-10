# Processed data — generated, do not edit by hand

This folder holds the cleaned dataset the visualization actually reads,
produced from the `.xlsx` files in `../raw/` by the processing script.

Expected output: a single `schools.json` in **long format** (one record per
school per year):

```json
[
  {
    "school_id": "040123",
    "school_name": "Lincoln Elementary",
    "year": 2019,
    "grade": "B",
    "ela_pass_rate": 72.4,
    "math_pass_rate": 68.1,
    "enrollment": 512
  }
]
```

Because everything here is regenerated from `raw/`, it's safe to delete and
rebuild at any time.

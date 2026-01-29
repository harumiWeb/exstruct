# RUB Report

This report summarizes Reconstruction Utility Benchmark (RUB) results.
Scores are computed on Stage B task accuracy using Markdown-only inputs.

## Summary by method

| method    |       rus |   avg_in |   avg_cost |   n |   partial_precision |   partial_recall |   partial_f1 |
|:----------|----------:|---------:|-----------:|----:|--------------------:|-----------------:|-------------:|
| exstruct  | 0.166667  |  844.167 | 0.00254708 |  12 |            0.666667 |         0.365278 |     0.436772 |
| html      | 0.0833333 |  994.333 | 0.00291833 |  12 |            0.583333 |         0.335317 |     0.399903 |
| image_vlm | 0         |  743.167 | 0.00231458 |  12 |            0.666667 |         0.343254 |     0.423112 |
| openpyxl  | 0.166667  |  769.083 | 0.00243771 |  12 |            0.5      |         0.37004  |     0.408236 |
| pdf       | 0         |  924.5   | 0.00265208 |  12 |            0.583333 |         0.252976 |     0.325493 |

## Summary by track

| track           | method    |       rus |   avg_in |   avg_cost |   n |   partial_precision |   partial_recall |   partial_f1 |
|:----------------|:----------|----------:|---------:|-----------:|----:|--------------------:|-----------------:|-------------:|
| structure_query | exstruct  | 0.166667  |  844.167 | 0.00254708 |  12 |            0.666667 |         0.365278 |     0.436772 |
| structure_query | html      | 0.0833333 |  994.333 | 0.00291833 |  12 |            0.583333 |         0.335317 |     0.399903 |
| structure_query | image_vlm | 0         |  743.167 | 0.00231458 |  12 |            0.666667 |         0.343254 |     0.423112 |
| structure_query | openpyxl  | 0.166667  |  769.083 | 0.00243771 |  12 |            0.5      |         0.37004  |     0.408236 |
| structure_query | pdf       | 0         |  924.5   | 0.00265208 |  12 |            0.583333 |         0.252976 |     0.325493 |

# RamerDouglasPeucker-ExcelVBA

An optimised Excel VBA implementation of the [Ramer–Douglas–Peucker](https://en.wikipedia.org/wiki/Ramer%E2%80%93Douglas%E2%80%93Peucker_algorithm) polyline simplification algorithm, with:

- **Two run modes**: fixed epsilon or goal-seek to a target point count
- **Pure-VBA toolbar** in the Add-ins tab (no third-party tools, no XML)
- **Iterative algorithm** for large datasets (no recursion stack risk)
- **Workbook generator** script to set up the required Excel structure automatically

---

## Quick Start

### 1. Generate the workbook template

```bash
pip install openpyxl
python create_workbook.py
```

This produces `RamerDouglasPeucker_Template.xlsx` with all tables, named ranges, and sample data pre-loaded.

### 2. Convert to macro-enabled workbook

1. Open `RamerDouglasPeucker_Template.xlsx` in Excel.
2. **File → Save As → Excel Macro-Enabled Workbook (.xlsm)** — give it any name you like.

### 3. Import the VBA module

1. Press **Alt+F11** to open the Visual Basic Editor.
2. **File → Import File** → select `RamerDouglasPeucker.bas`.

### 4. Add the workbook event hooks

In the VBE Project pane, double-click **ThisWorkbook** and paste:

```vba
Private Sub Workbook_Open()
    CreateRDPToolbar
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DeleteRDPToolbar
End Sub
```

### 5. Save, close, and reopen

The **RDP Simplify** toolbar now appears in the **Add-ins** tab every time the workbook is opened.

---

## Workbook Structure

| Location | Name | Purpose |
|---|---|---|
| Sheet1, columns A–B | `Table1` | Input X/Y coordinates |
| Sheet1, cell E1 | named range `epsilon` | Simplification tolerance |
| Sheet1, cell E2 | named range `rdpTargetCount` | Target output point count (goal-seek mode) |
| Sheet2, columns A–B | `Table2` | Simplified output coordinates |

The script pre-fills `epsilon = 5` and `rdpTargetCount = 20`. Edit these cells directly before clicking toolbar buttons.

---

## Toolbar Buttons

| Button | Action |
|---|---|
| **Run Simplification** | Simplifies Table1 using the current `epsilon` value, writes result to Table2 |
| **Goal Seek** | Binary-searches for the epsilon that produces output closest to `rdpTargetCount` points, writes the found epsilon back to E1, then runs |
| **Clear Output** | Deletes all data rows from Table2 |

---

## Choosing Epsilon

Epsilon is the maximum perpendicular distance (in the same units as your coordinates) that a point may deviate from the simplified line before it is kept.

For the included sample data (pixel coordinates, range ~400 × 580):

| Epsilon | Approximate output points (from 197) |
|---|---|
| 1 | ~120 |
| 5 | ~35 |
| 10 | ~20 |
| 20 | ~12 |

Use **Goal Seek** when you want a specific count and don't want to tune epsilon manually.

---

## How Goal Seek Works

There is no closed-form formula mapping a target point count to an epsilon value — the output count is a monotonically decreasing step function of epsilon with N−2 steps, and each step depends on the geometry of the curve.

What *is* analytically computable is a tight upper bound: the maximum perpendicular distance of any point from the baseline chord (first point → last point), computed in a single O(N) pass. Goal Seek binary-searches the range `[0, upper_bound]` in at most 50 iterations, which is sufficient to converge to within 10⁻⁹ regardless of dataset size. The found epsilon is written back to cell E1 so you can reuse it with **Run Simplification**.

---

## Algorithm Notes

### Optimisations over the original implementation

| Area | Original | Optimised |
|---|---|---|
| Distance formula | `Sqr()` called on every point, every call | Squared distances compared — zero `Sqr()` in the hot path |
| Denominator | Recomputed inside the inner loop | Hoisted outside the loop (one computation per recursive call) |
| Large datasets | Recursive only (stack overflow risk above ~10 000 pts) | Iterative path (explicit stack) used automatically for > 5 000 points |
| Array output | Cell-by-cell loop (2×N Excel boundary crossings) | Single bulk `Range.Value = array` assignment |
| Variable types | `Integer` (overflows at 32 767) | `Long` throughout |
| Code quality | No `Option Explicit`, undeclared variables | `Option Explicit`, all variables declared, error handling on all entry points |

### Squared-distance equivalence

The standard perpendicular distance is:

```
d = |cross_product| / sqrt(line_length²)
```

Comparing `d > epsilon` is equivalent to comparing `cross² > epsilon² × line_length²`, which avoids the square root entirely and is safe whenever `line_length² > 0` (the degenerate case is explicitly handled).

---

## Test Data

`data.csv` contains 197 X/Y coordinate pairs (pixel coordinates of a curved line) suitable for testing the algorithm.

---

## License

MIT — see [LICENSE](LICENSE).

# Module 2 — Excel Fundamentals and Functions

## Learning outcomes

Students will be able to:

- identify workbooks, worksheets, rows, columns, cells, ranges, and the formula bar;
- distinguish numbers, text, dates, logical values, blanks, and errors;
- create simple formulas;
- use relative, absolute, and mixed references;
- use common built-in Excel functions; and
- explain zero, errors, and a simple iterative process.

## 100-minute sequence

| Minutes | Activity |
|---:|---|
| 0–10 | Interface, cell addresses, and data types |
| 10–28 | Formulas and relative/absolute references |
| 28–45 | Volume-and-cost table demonstration |
| 45–60 | `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, and `ROUND` |
| 60–72 | Zero, blank cells, text, and error values |
| 72–84 | A simple iterative calculation |
| 84–97 | Three-question quiz |
| 97–100 | Review and checklist |

**Minimum product:** one volume-and-cost table whose formulas copy correctly, a function summary, and a five-step iteration table.

## 1. Main Excel components

- **Workbook:** an Excel file.
- **Worksheet:** a sheet inside a workbook.
- **Cell:** an intersection of a row and column, such as `B3`.
- **Range:** a group of cells, such as `B3:D10`.
- **Formula bar:** displays the cell's actual content.
- **Name box:** displays or accepts a cell/range address.

A cell contains a value or formula. Formatting changes its display, not its underlying value. For example, `0.15` displayed as `15%` remains numerically `0.15`.

## 2. Data types

| Content | Type | Note |
|---|---|---|
| `12.5` | number | can be calculated |
| `Channel A` | text | label or code |
| `26/08/2026` | date | stored as a date serial number |
| `TRUE` | logical | true/false |
| an unfilled cell | blank | different from zero |
| `#DIV/0!` | error | invalid operation, such as division by zero |

Do not enter a value as `12 m`. Store `12` as a number and put `(m)` in the column heading.

## 3. Simple formulas

Every formula begins with `=`.

```excel
=B2*C2
=(B2+C2)/2
=PI()*B2^2/4
```

Use parentheses to communicate the intended order of operations.

## 4. Cell references

Assume column B contains volumes and cell `F1` contains a unit price.

| Formula | Reference type | After copying down one row |
|---|---|---|
| `=B2*F1` | relative | `=B3*F2` |
| `=B2*$F$1` | F1 absolute | `=B3*$F$1` |
| `=$B2*F$1` | mixed | column B and row 1 remain fixed |

While editing a cell reference, press **F4** to cycle through `F1`, `$F$1`, `F$1`, and `$F1`. F5 opens *Go To*; it does not toggle absolute references.

## 5. Practice — concrete volume and cost

Create this table:

| A | B | C | D | E |
|---|---:|---:|---:|---:|
| Segment | Length (m) | Width (m) | Height (m) | Volume (m³) |
| S1 | 10 | 0.2 | 0.3 | formula |
| S2 | 12 | 0.2 | 0.3 | formula |

Place the concrete unit price in `H2`. In column F calculate cost:

```excel
=E2*$H$2
```

Copy the formula down. Confirm that the volume address changes while `H2` remains fixed.

## 6. Built-in functions

Functions accept inputs, or **arguments**, and return a result. Build a summary for `E2:E11`:

```excel
=SUM(E2:E11)
=AVERAGE(E2:E11)
=MIN(E2:E11)
=MAX(E2:E11)
=COUNT(E2:E11)
=ROUND(AVERAGE(E2:E11),3)
```

`COUNT` counts numeric cells; `COUNTA` counts nonblank cells. The distinction matters when a dataset contains text or labels.

## 7. Zero, blanks, and errors

- Zero is a valid number, but may make division undefined.
- A blank cell means data has not been supplied; it should not always be treated as zero.
- Text that looks like a number may disrupt calculations.
- `#DIV/0!`, `#VALUE!`, `#NAME?`, `#REF!`, and `#N/A` provide clues about the problem.

Do not hide every error immediately. Identify its cause. Use `IFERROR` only after deliberately defining the replacement behaviour.

## 8. Iteration example

Iteration repeats a rule until the result is sufficiently stable. To approximate `√2`, use:

```text
x_new = (x_old + 2/x_old) / 2
```

Start at `x₀ = 1`:

| Iteration | old x | new x | absolute change |
|---:|---:|---:|---:|
| 1 | 1 | `=(B2+2/B2)/2` | `=ABS(C2-B2)` |

The new value becomes the old value on the next row. Stop when the change is smaller than a tolerance such as `0.000001`. Avoid enabling workbook iteration before understanding circular references and stopping conditions.

## High-impact quiz — 3 questions

### 1. Predict

Cell `C2` contains `=A2*$B$1`. What does it become when copied to `C5`? What would happen without the dollar signs?

### 2. Fix

The velocity formula `=B2/C2` returns `#DIV/0!`. Write a check that distinguishes a blank `C2`, `C2=0`, and `C2>0`. Do not simply mask every case with `IFERROR`.

### 3. Explain

Starting from `x₀=1`, calculate the first two iterations for the square root of two. Explain why the process needs a tolerance and stopping rule.

<details>
<summary>Answer key and evidence of understanding</summary>

1. It becomes `=A5*$B$1`. Without `$`, B1 also shifts, becoming B4.
2. Example: `=IF(C2="","Area is missing",IF(C2=0,"Area cannot be zero",B2/C2))`.
3. `x₁=1.5`; `x₂≈1.4166667`. The tolerance defines when change is small enough; without a stopping condition, iteration may continue indefinitely or create an uncontrolled circular reference.

</details>

## Checklist

- [ ] I distinguish values, formatting, and formulas.
- [ ] I choose relative or absolute references deliberately.
- [ ] I distinguish zero, blank, text, and errors.
- [ ] I can state an iterative rule and its stopping condition.

[← Module 1](01-computers-and-operating-systems.md) · [Module list](README.md) · [Module 3 →](03-excel-if.md)

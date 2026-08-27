# Module 3 — Decision-Making with Excel IF

## Learning outcomes

Students will be able to:

- apply Excel functions within a tabular calculation;
- write single and nested `IF` formulas;
- combine `IF` with `AND` and `OR`;
- classify student marks into A–E categories; and
- test a formula exactly at every decision boundary.

## 100-minute sequence

| Minutes | Activity |
|---:|---|
| 0–10 | Inspect a grade table and identify its decisions |
| 10–25 | Review functions and calculate a weighted mark |
| 25–40 | Single `IF` and comparison operators |
| 40–55 | Nested `IF` and condition order |
| 55–75 | Build an A–E grade classifier |
| 75–84 | Test boundaries and invalid inputs |
| 84–97 | Three-question short exercise |
| 97–100 | Checklist |

**Minimum product:** a student-grade table that calculates a weighted result, pass status, and A–E grade correctly at every boundary.

## 1. Single IF

The general structure is:

```excel
=IF(condition,value_if_true,value_if_false)
```

For a pass threshold of 60:

```excel
=IF(B2>=60,"Pass","Not yet passed")
```

The `>=` is important. With only `>`, a mark of exactly 60 would be placed in the wrong category.

## 2. Combining conditions

A mark is valid only when it lies between 0 and 100:

```excel
=IF(AND(B2>=0,B2<=100),"Valid","Invalid")
```

A student requires remedial work when the mark is below 60 **or** attendance is below 75%:

```excel
=IF(OR(B2<60,C2<75%),"Remedial","No remedial")
```

Use `AND` when every condition must be true. Use `OR` when any one condition is sufficient.

## 3. Nested IF for A–E grades

Use these practice rules:

| Range | Grade |
|---|---|
| 80–100 | A |
| 70–<80 | B |
| 60–<70 | C |
| 50–<60 | D |
| 0–<50 | E |

Validate the overall domain before classifying:

```excel
=IF(OR(B2<0,B2>100),"Invalid",
 IF(B2>=80,"A",
 IF(B2>=70,"B",
 IF(B2>=60,"C",
 IF(B2>=50,"D","E")))))
```

Where supported, `IFS` can improve readability:

```excel
=IF(OR(B2<0,B2>100),"Invalid",
 IFS(B2>=80,"A",B2>=70,"B",B2>=60,"C",B2>=50,"D",TRUE,"E"))
```

Conditions are checked from the highest boundary downward. Once a true condition is found, later conditions are not used.

## 4. Practice — student-grade table

Create these columns:

| Column | Content |
|---|---|
| A | Student name |
| B–D | Assignment, midterm, final exam |
| E | Final mark |
| F | Validity status |
| G | Letter grade |

Use the practice weighting `30% assignment + 30% midterm + 40% final exam`:

```excel
=B2*30%+C2*30%+D2*40%
```

Then classify the result. Test at least these final marks:

```text
-1, 0, 49, 50, 59, 60, 69, 70, 79, 80, 100, 101
```

> These categories are learning data, not an official academic grading policy.

## 5. Common errors

- excluding an exact boundary, for example using `>80`;
- starting conditions at the lowest threshold, so every higher mark enters the first category;
- storing numbers as text;
- assigning a letter grade to a value outside 0–100; and
- using a long nested formula without documenting its rules.

For rules that change often, a lookup table is usually easier to maintain than deeply nested IF statements. Treat that method as an extension.

## 6. Boundary-test table

Before trusting the formula, write expected and actual outcomes:

| Input | Expected | Actual | Pass? |
|---:|---|---|---|
| 49 | E | fill in | fill in |
| 50 | D | fill in | fill in |
| 59 | D | fill in | fill in |
| 60 | C | fill in | fill in |
| 79 | B | fill in | fill in |
| 80 | A | fill in | fill in |
| 101 | Invalid | fill in | fill in |

Testing one ordinary value does not prove that all boundaries are correct.

## Short exercise — 3 questions

### 1. Predict

What is the result of this formula for `B2=80`, and why?

```excel
=IF(B2>80,"A",IF(B2>70,"B","C"))
```

### 2. Fix

Repair the formula so A begins at `>=80`, B at `>=70`, and any input outside 0–100 returns `Invalid`.

### 3. Explain

A student receives 100 for assignments, 50 for the midterm, and 50 for the final exam. Predict the final mark and grade. Explain why testing only a mark of 75 is insufficient.

<details>
<summary>Answer key and evidence of understanding</summary>

1. The result is B because `80>80` is false and `80>70` is true.
2. `=IF(OR(B2<0,B2>100),"Invalid",IF(B2>=80,"A",IF(B2>=70,"B","C")))`.
3. Final mark: `30+15+20=65`, therefore C. A test at 75 covers only a middle branch; every exact boundary, domain endpoint, and invalid value also needs a test.

</details>

## Checklist

- [ ] I can write single and nested IF formulas.
- [ ] I deliberately choose `>` or `>=`.
- [ ] I validate the domain before classifying.
- [ ] I test exactly at every boundary.

## Further reading

1. Michael Alexander & Dick Kusleika, *Microsoft Excel 365 Bible*, 2nd ed., Wiley, 2025.
2. Wayne L. Winston, *Microsoft Excel Data Analysis and Business Modeling*, 6th ed., Microsoft Press, 2019.
3. Bernard Liengme & Keith Hekman, *Liengme’s Guide to Excel 2016 for Scientists and Engineers*, Academic Press, 2019.
4. E. Joseph Billo, *Excel for Scientists and Engineers: Numerical Methods*, Wiley, 2007.

[← Module 2](02-excel-fundamentals.md) · [Module list](README.md) · [Module 4 →](04-algorithms.md)

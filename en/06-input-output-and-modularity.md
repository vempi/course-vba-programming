# Module 6 — Input/Output, Variables, Arrays, Function, and Subroutine

## Learning outcomes

Students will be able to:

- read and write data with `Cells`, `Range`, `InputBox`, and `MsgBox`;
- declare variables with appropriate types and scope;
- use an array for a collection of values;
- distinguish a `Function` from a `Sub`; and
- build a modular program that processes a table.

## 100-minute sequence

| Minutes | Activity |
|---:|---|
| 0–10 | Identify IPO in the previous linear macro |
| 10–25 | `Cells`, `Range`, `InputBox`, and `MsgBox` |
| 25–40 | Variable types and local/public scope |
| 40–55 | One- and two-dimensional arrays |
| 55–70 | Difference between `Function` and `Sub` |
| 70–85 | Process a volume table |
| 85–97 | Three-question quiz |
| 97–100 | Compile and complete checklist |

**Minimum product:** one volume function and one `Sub` that reads at least five rows into an array, processes them, and writes results to Excel.

## 1. Simple input and output

### `Range` and `Cells`

```vb
Dim value As Double

value = Range("B2").Value
Cells(2, 3).Value = value * 2
```

`Range("B2")` is readable for a fixed address. `Cells(row, column)` is useful when the row or column changes inside a loop.

Prefer an explicit worksheet:

```vb
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("VolumeData")

value = ws.Range("B2").Value
ws.Cells(2, "F").Value = value
```

### `InputBox` and `MsgBox`

```vb
Sub InputOutputDemo()
    Dim projectName As String

    projectName = InputBox("Enter the project name:")

    If projectName = "" Then
        MsgBox "No project name was entered.", vbExclamation
    Else
        MsgBox "Project: " & projectName, vbInformation
    End If
End Sub
```

A standard `InputBox` returns text. Use `IsNumeric` before converting numeric text with `CDbl`.

## 2. Variables and data types

```vb
Dim length_m As Double
Dim segmentCount As Long
Dim segmentName As String
Dim dataIsValid As Boolean
Dim surveyDate As Date
```

Use:

- `Double` for fractional engineering quantities;
- `Long` for counts and row indices;
- `String` for names and codes;
- `Boolean` for true/false state; and
- `Date` for dates.

`Option Explicit` requires declarations and exposes misspelled variable names.

## 3. Local and public scope

A local variable is declared inside a procedure and is available only there:

```vb
Sub LocalExample()
    Dim total_m3 As Double
    total_m3 = 10
End Sub
```

A public variable is declared at the top of a standard module:

```vb
Option Explicit
Public activeProjectName As String
```

Prefer local variables. A public variable is justified only when state must genuinely be shared among modules or forms; its value is harder to trace because several locations may change it.

## 4. Arrays

An array stores several values of the same kind under one name.

```vb
Sub ArrayDemo()
    Dim volumes_m3(1 To 3) As Double
    Dim i As Long
    Dim total_m3 As Double

    volumes_m3(1) = 2.5
    volumes_m3(2) = 3.5
    volumes_m3(3) = 4

    For i = LBound(volumes_m3) To UBound(volumes_m3)
        total_m3 = total_m3 + volumes_m3(i)
    Next i

    MsgBox "Total = " & total_m3 & " m³"
End Sub
```

An Excel block can be loaded directly into a two-dimensional `Variant` array:

```vb
Dim data As Variant
data = ws.Range("B2:D6").Value

' data(1,1) is B2; data(1,2) is C2; data(1,3) is D2
```

## 5. Function and Subroutine

A `Function` accepts parameters and returns a value:

```vb
Public Function RectangularVolume( _
    ByVal length_m As Double, _
    ByVal width_m As Double, _
    ByVal height_m As Double) As Double

    If length_m <= 0 Or width_m <= 0 Or height_m <= 0 Then
        RectangularVolume = -1
    Else
        RectangularVolume = length_m * width_m * height_m
    End If
End Function
```

A `Sub` coordinates actions: reading a table, calling the function, and writing output.

```vb
Public Sub CalculateVolumeTable()
    Dim ws As Worksheet
    Dim data As Variant
    Dim results() As Variant
    Dim i As Long
    Dim volume_m3 As Double

    Set ws = ThisWorkbook.Worksheets("VolumeData")
    data = ws.Range("B2:D6").Value
    ReDim results(1 To UBound(data, 1), 1 To 2)

    For i = 1 To UBound(data, 1)
        If IsNumeric(data(i, 1)) And IsNumeric(data(i, 2)) And _
           IsNumeric(data(i, 3)) Then

            volume_m3 = RectangularVolume(CDbl(data(i, 1)), _
                                          CDbl(data(i, 2)), _
                                          CDbl(data(i, 3)))

            If volume_m3 >= 0 Then
                results(i, 1) = volume_m3
                results(i, 2) = "OK"
            Else
                results(i, 1) = Empty
                results(i, 2) = "Invalid dimensions"
            End If
        Else
            results(i, 1) = Empty
            results(i, 2) = "Not numeric"
        End If
    Next i

    ws.Range("E2:F6").Value = results
    ws.Range("E2:E6").NumberFormat = "0.000"
End Sub
```

### `VolumeData` worksheet

| Column | Content |
|---|---|
| A | Segment |
| B | Length (m) |
| C | Width (m) |
| D | Height (m) |
| E | Volume (m³) |
| F | Status |

Include five rows, with one zero dimension and one text value, so validation can be observed.

## 6. Why modular design matters

`RectangularVolume` knows nothing about cell locations. It is easy to test and can be reused by the table macro and by the UserForm in Module 7. `CalculateVolumeTable` coordinates Excel input/output but does not own the geometry formula.

Test the function in the Immediate Window:

```text
? RectangularVolume(10, 2, 0.3)
? RectangularVolume(0, 2, 0.3)
```

Expected outputs are `6` and `-1`.

## High-impact quiz — 3 questions

### 1. Predict

Variable `total_m3` is declared inside `Sub A`. Can `Sub B` read it directly? What redesign is safer than making the variable public?

### 2. Fix

The third row contains the text `two`. Identify the condition that prevents `CDbl("two")` from running, and explain what the program writes as output.

### 3. Explain

Why should the volume formula belong to `RectangularVolume`, while worksheet reading/writing belongs to `CalculateVolumeTable`? State one testing benefit.

<details>
<summary>Answer key and evidence of understanding</summary>

1. No. A local variable exists only in its procedure. Pass the value as a parameter or return it from a Function.
2. `IsNumeric` is checked before `CDbl`; the volume is blank and status becomes `Not numeric`.
3. Separating responsibilities keeps the calculation independent of the worksheet. The function can be tested directly with known inputs and expected outputs.

</details>

## Checklist

- [ ] Every variable is declared with an appropriate type.
- [ ] Input is checked before conversion.
- [ ] I distinguish local and public scope.
- [ ] I distinguish the roles of Function and Sub.
- [ ] All five rows are processed even when one row is invalid.

[← Module 5](05-vba-and-linear-macros.md) · [Module list](README.md) · [Module 7 →](07-userforms-and-controls.md)

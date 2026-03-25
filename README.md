# FDC Calculation Automation Tool

Automating artillery azimuth and distance calculations using Excel VBA.

---
# Operational Context

During my military service, I worked in an artillery unit operating mortar systems during training exercises.

<img width="353" height="452" alt="Screenshot 2026-03-25 at 3 18 07 AM" src="https://github.com/user-attachments/assets/3dbc9bcf-37ba-4154-8638-3721fc272cd1" />


In mortar operations, accurate firing requires calculating two key values before each shot:

- **Azimuth** – the direction of fire
- **Distance** – the distance to the target

These calculations were traditionally performed manually using a plotting tool.

---
# Overview

During artillery operations, firing data such as **azimuth** and **range** were traditionally calculated manually using the **M17 plotting board**.

The manual process involved determining the coordinates of the artillery position and the target, calculating the coordinate difference \((X, Y)\), plotting the point on the M17 board, rotating the board until the point aligned with the vertical axis, and then reading the azimuth and distance directly from the board.

Although this method worked, it required several manual steps and relied on visual estimation, which could introduce small errors.

During training exercises, we often had to process **a large number of target coordinates within a short time window**, usually under significant time pressure. Because these exercises simulate real combat conditions, both **speed and accuracy are critical**.

To improve this workflow, I analyzed the underlying geometry of the manual process and derived a mathematical model that could produce the same results directly.

I then implemented the calculation using **Excel VBA**, allowing azimuth and distance to be computed instantly from coordinate inputs.

---

# Original Manual Process (M17 Plotting Board)

Before automation, artillery firing data was calculated using the **M17 plotting board**.

The steps were:

1. Determine the coordinates of the artillery position.
2. Obtain the target coordinates from the map.
3. Compute the coordinate difference to obtain **X** and **Y**.
4. Plot the \((X, Y)\) point on the M17 board.
5. Rotate the board so that the plotted point aligns with the vertical axis.
6. Read the **azimuth** and **distance** directly from the board scale.

When using the M17 plotting board, coordinates are often scaled by a ratio (e.g., 1:2) so that the point fits within the board. The final distance reading must then be multiplied by the same ratio.

This manual process required careful plotting and interpretation, and small reading errors could affect the result.

### M17 Plotting Board

<img width="602" height="271" alt="Screenshot 2026-03-25 at 1 56 10 AM" src="https://github.com/user-attachments/assets/3e220390-5176-44fe-98b1-6f76095670df" />


### Example of Manual Reading

<img width="597" height="276" alt="Screenshot 2026-03-25 at 1 56 24 AM" src="https://github.com/user-attachments/assets/d6d3a6ad-e4a1-4743-b452-7bd1039ced3c" />

<img width="601" height="268" alt="Screenshot 2026-03-25 at 1 56 36 AM" src="https://github.com/user-attachments/assets/71f3b197-6eb3-4384-a779-bf605ba342f4" />

---

# Mathematical Model

After analyzing the geometry behind the plotting method, the following formulas were derived to compute the firing data directly.

### Distance

Distance = √(X² + Y²) × 10

### Azimuth

Azimuth = atan2(X, Y) × 3200 / π + 1600

The additional **1600 mil offset** aligns the mathematical coordinate system with the reading convention used by the **M17 plotting board**.

---

# Excel VBA Implementation

The formulas were implemented in **Excel VBA**, allowing users to input coordinate values and calculate the result instantly.

### VBA Code

```vba
Sub CalculateFDC()

    Dim X As Double
    Dim Y As Double
    Dim Azimuth As Double
    Dim Distance As Double
    Dim PiValue As Double

    PiValue = WorksheetFunction.Pi()

    X = Range("B4").Value
    Y = Range("B5").Value

    Distance = Sqr((X ^ 2) + (Y ^ 2)) * 10

    Azimuth = WorksheetFunction.Atan2(X, Y) * 3200 / PiValue + 1600

    If Azimuth < 0 Then
        Azimuth = Azimuth + 6400
    ElseIf Azimuth >= 6400 Then
        Azimuth = Azimuth - 6400
    End If

    Range("B8").Value = Distance
    Range("B9").Value = Azimuth

End Sub
```
This tool automates the original manual calculation process and provides results immediately after entering the coordinate values.

---

## Example

### Input

- **X = -60**
- **Y = 40**

### Output

- **Distance ≈ 721.11**
- **Azimuth ≈ 4201.07 mil**

### Excel Tool Interface

<img width="289" height="202" alt="Screenshot 2026-03-25 at 2 03 27 AM" src="https://github.com/user-attachments/assets/7f9802d7-1acc-4cfa-b496-f66b69f5a206" />

---

## Before vs After

### Before

- Manual plotting on M17 board
- Board rotation required
- Visual reading of azimuth and distance
- Small reading errors possible
- More time required

### After

- Direct coordinate input in Excel
- Automatic trigonometric calculation
- Instant output of azimuth and distance
- Improved consistency
- Reduced manual effort

---

## Key Features

- Automates manual artillery calculation workflow
- Computes **azimuth** and **distance** directly from coordinate inputs
- Reduces manual plotting and reading errors
- Improves consistency in repeated calculations
- Demonstrates process automation using **Excel VBA**

---

## Technologies Used

- **Excel VBA**
- **Trigonometry**
- **Mathematical modeling**
- **Process automation**

---

## Skills Demonstrated

- Translating a manual workflow into a mathematical model
- Automating repetitive calculations using Excel VBA
- Applying trigonometric formulas to a real-world operational problem
- Improving efficiency and reducing human error in a manual system

---

## Project Structure

```text
fdc-calculation-automation/
│
├── README.md
├── images/
│   ├── m17-board.png
│   ├── m17-rotation.png
│   └── excel-tool.png
└── fdc_calculation_tool.xlsm

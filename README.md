# Sales Dashboard with Dynamic Slicer Control

An interactive Excel dashboard built using Pivot Tables, Pivot Charts, Slicers, and VBA. This project enables users to selectively connect dashboard components to a Region slicer using checkboxes, providing flexible and dynamic data analysis.

Developed by **Muhammad Waleed** ([winx007](https://github.com/winx007)).

---

## 🚀 Features

*   **Interactive Region Slicer**: Filters multiple visualizations simultaneously.
*   **Multiple Pivot Table Dashboards**: Separated modules for distinct analytical insights.
*   **Checkbox-Controlled Slicer Connections**: Connect or disconnect dashboard components dynamically using checkboxes.
*   **Real-time VBA Automation**: Leverages VBA to toggle Slicer Cache connections instantly.
*   **Pivot Charts**: Visual representations (bar, line, and pie charts) for data-driven decisions.

---

## 📊 Dashboard Modules

### 🔹 Dashboard #01: Sales Executive Performance
*   Total Sales Table
*   Bar Chart Analysis

### 🔹 Dashboard #02: Sales Comparison
*   Sales Executive Performance Comparison Dashboard

### 🔹 Dashboard #03: Target Achievement
*   Target Achievement Analysis
*   Pie Chart Visualization

### 🔹 Dashboard #04: Trend Analysis
*   Away From Target Analysis
*   Trend Analysis using Line Chart

---

## 🛠️ How It Works

Each dashboard module has a dedicated checkbox linked to a worksheet cell.

| Checkbox | Pivot Table | Linked Cell |
| :--- | :--- | :--- |
| **Dashboard #01** | `PivotTable1` | `A1` |
| **Dashboard #02** | `PivotTable2` | `D1` |
| **Dashboard #03** | `PivotTable3` | `G1` |
| **Dashboard #04** | `PivotTable4` | `J1` |

*   **When Checked (`True`)**: The corresponding Pivot Table is connected to the Region slicer, updating dynamically.
*   **When Unchecked (`False`)**: The Pivot Table is disconnected from the slicer, freezing its state so it is unaffected by slicer changes.

---

## 💻 VBA Automation Setup

The repository contains `Macro.vba`, which automates slicer connections in the background. 

### 1. The Core Macro (`Macro.vba`)
Add this code into a **Standard Module** (`Insert` > `Module`) in the VBA Editor:

```vba
Sub new_init()
    Dim sc As SlicerCache
    Set sc = ActiveWorkbook.SlicerCaches("Slicer_Region")

    On Error Resume Next

    ' Dashboard 1
    If Sheet1.Range("A1").Value Then
        sc.PivotTables.AddPivotTable Sheet1.PivotTables("PivotTable1")
    Else
        sc.PivotTables.RemovePivotTable Sheet1.PivotTables("PivotTable1")
    End If

    ' Dashboard 2
    If Sheet1.Range("D1").Value Then
        sc.PivotTables.AddPivotTable Sheet1.PivotTables("PivotTable2")
    Else
        sc.PivotTables.RemovePivotTable Sheet1.PivotTables("PivotTable2")
    End If

    ' Dashboard 3
    If Sheet1.Range("G1").Value Then
        sc.PivotTables.AddPivotTable Sheet1.PivotTables("PivotTable3")
    Else
        sc.PivotTables.RemovePivotTable Sheet1.PivotTables("PivotTable3")
    End If

    ' Dashboard 4
    If Sheet1.Range("J1").Value Then
        sc.PivotTables.AddPivotTable Sheet1.PivotTables("PivotTable4")
    Else
        sc.PivotTables.RemovePivotTable Sheet1.PivotTables("PivotTable4")
    End If

    On Error GoTo 0
End Sub
```

### 2. Event-Driven Execution (Auto-Run)
To run this macro automatically the instant a checkbox is clicked without manual execution, copy this event-driven sub into the **Worksheet Code** of your dashboard sheet:

```vba
Private Sub Worksheet_Calculate()
    ' Automatically runs the macro whenever the linked cells calculate/change
    Call new_init
End Sub
```

> [!IMPORTANT]
> **File Format Notice:** 
> When saving your dashboard workbook, make sure to save it as **Excel Macro-Enabled Workbook (`.xlsm`)**. If saved as `.xlsx`, Excel will automatically strip out all VBA macros.

---

## 📁 Project Structure

```text
Sales-Dashboard/
│
├── Project File.xlsm          # Main spreadsheet template
├── Macro.vba                  # VBA macro source code
├── README.md                  # Project documentation
└── Screenshots/
    └── dashboard.png          # Dashboard layout preview
```

---

## 📈 Screenshot Preview

![Dashboard Preview](Screenshots/dashboard.png)

---

## 👤 Author

**Muhammad Waleed**
*   Excel Dashboard Development • VBA Automation • Data Analytics
*   GitHub: [@winx007](https://github.com/winx007)

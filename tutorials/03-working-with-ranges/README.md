# Tutorial 3: Working with Ranges

> Master reading, writing, formatting, and manipulating cell ranges

## 🎯 Learning Objectives

By the end of this tutorial, you will be able to:
- Read and write data to ranges of any size
- Apply cell formatting (bold, colors, borders)
- Work with formulas across ranges
- Use find and replace operations
- Copy and move data between ranges

## ⏱️ Time Required

Approximately 15-20 minutes

## 📋 Prerequisites

- Completed [Tutorial 2: Your First Excel Automation](../02-first-automation/README.md)

## 📁 Files Used

| File | Description |
|------|-------------|
| [sample-data.xlsx](files/sample-data.xlsx) | Sample sales data for practice |

> 💡 Download the sample file, or create your own following Step 1.

## 🚀 Steps

### Step 1: Create Sample Data

First, let's create a workbook with some data to work with:

**Prompt:**
```
Create C:\Temp\sales-data.xlsx with this sales data starting at A1:

Product, Q1, Q2, Q3, Q4
Widgets, 1500, 1800, 2100, 2400
Gadgets, 2200, 2000, 2500, 2800
Gizmos, 1800, 1900, 2200, 2100

Save and close.
```

---

### Step 2: Read a Specific Range

Let's read just the Q1 and Q2 data:

**Prompt:**
```
Open C:\Temp\sales-data.xlsx and read the range B1:C4
```

**Expected result:**
```
Q1    | Q2
1500  | 1800
2200  | 2000
1800  | 1900
```

---

### Step 3: Apply Basic Formatting

Make our headers stand out:

**Prompt:**
```
Open C:\Temp\sales-data.xlsx and format the header row A1:E1 with:
- Bold text
- Blue background color (#4472C4)
- White font color
Then save and close.
```

**Expected result:**
The header row now has a professional blue background with white bold text.

---

### Step 4: Add Conditional Formatting (Coming Soon)

> 📝 **Note:** Conditional formatting support is planned for a future release.

For now, you can apply static formatting based on values manually:

**Prompt:**
```
Open C:\Temp\sales-data.xlsx and highlight any cells in range B2:E4 that contain values over 2200 with a green background color (#C6EFCE). Save and close.
```

---

### Step 5: Add Formulas to Calculate Totals

Let's add row and column totals:

**Prompt:**
```
Open C:\Temp\sales-data.xlsx and:
1. Add "Total" in cell F1
2. In F2, add a formula to sum B2:E2 (Widget totals)
3. In F3, add a formula to sum B3:E3 (Gadget totals)
4. In F4, add a formula to sum B4:E4 (Gizmo totals)
5. Add "Total" in A5
6. In B5, add a formula to sum B2:B4 (Q1 total)
7. In C5, D5, E5, and F5, add similar sum formulas for each column
Save and close.
```

**Expected result:**
Your spreadsheet now shows totals for each row and column, with a grand total in F5.

---

### Step 6: Find and Replace

Let's rename a product:

**Prompt:**
```
Open C:\Temp\sales-data.xlsx and replace all occurrences of "Widgets" with "Super Widgets". Save and close.
```

**Expected result:**
Cell A2 now shows "Super Widgets" instead of "Widgets".

---

### Step 7: Copy Data Between Ranges

Let's copy Q4 data to a new location:

**Prompt:**
```
Open C:\Temp\sales-data.xlsx and:
1. Copy the range E1:E4 (Q4 data with header)
2. Paste it to column G starting at G1
3. Change the header G1 to "Q4 Copy"
Save and close.
```

**Expected result:**
Column G now contains a copy of the Q4 data.

---

### Step 8: Number Formatting

Format numbers as currency or percentages:

**Prompt:**
```
Open C:\Temp\sales-data.xlsx and:
1. Format cells B2:F5 as currency with no decimal places ($#,##0)
2. Save and close
```

**Expected result:**
All numeric cells now display with dollar signs (e.g., $1,500, $2,200).

---

### Step 9: Working with Large Ranges

For bulk operations, specify the full range:

**Prompt:**
```
Read the used range from C:\Temp\sales-data.xlsx (get all data that has content)
```

**Expected result:**
Returns all data in the worksheet, automatically detecting the boundaries.

---

### Step 10: Clear Content vs Clear All

Understand the difference:

**Prompt:**
```
Open C:\Temp\sales-data.xlsx and:
1. Clear the contents of G1:G4 (removes values but keeps formatting)
2. Save and close
```

To also remove formatting:

**Prompt:**
```
Open C:\Temp\sales-data.xlsx and clear all (contents AND formatting) from G1:G4, then save and close.
```

---

## ✅ Summary

You've mastered these range operations:

| Operation | Tool Action | Use Case |
|-----------|-------------|----------|
| Read range | `get-values` | Extract data from cells |
| Write range | `set-values` | Add or update data |
| Get formulas | `get-formulas` | See formulas, not results |
| Set formulas | `set-formulas` | Add calculations |
| Format | `format-range` | Colors, fonts, borders |
| Find | `find` | Search for content |
| Replace | `replace` | Find and update content |
| Copy | `copy` | Duplicate data |
| Clear contents | `clear-contents` | Remove values only |
| Clear all | `clear-all` | Remove values + formatting |

## 💡 Pro Tips

1. **Use 2D arrays for bulk writes** - Writing `[[1,2,3],[4,5,6]]` is faster than individual cells
2. **Get used range first** - Understand your data boundaries before operating
3. **Format after data entry** - Add formatting as a final step
4. **Use named ranges** - Create named ranges for frequently accessed areas

## 🔧 Troubleshooting

### Formatting not applying
- Check the range address is correct
- Some formats require specific data types (e.g., currency needs numbers)

### Formula errors (#REF!, #VALUE!)
- Verify the cell references exist
- Check for circular references

### Data overwritten unexpectedly
- Always verify target range is empty before writing
- Use `get-values` to check first

## ➡️ Next Steps

Continue your learning:

**[Tutorial 4: Excel Tables →](../04-excel-tables/README.md)**

Learn to work with structured Excel Tables for better data management!

---

## 🧪 Practice Exercises

1. **Exercise 1:** Create a grade book with students, test scores, and average formulas

2. **Exercise 2:** Format a range with alternating row colors (like a zebra stripe)

3. **Exercise 3:** Create a lookup table and use it with VLOOKUP formulas

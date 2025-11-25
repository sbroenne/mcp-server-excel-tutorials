# Tutorial 2: Your First Excel Automation

> Create, read, and modify Excel files using natural language

## 🎯 Learning Objectives

By the end of this tutorial, you will be able to:
- Create new Excel workbooks
- Write data to cells
- Read data from Excel files
- Save and close workbooks properly

## ⏱️ Time Required

Approximately 10-15 minutes

## 📋 Prerequisites

- Completed [Tutorial 1: Installation & Setup](../01-getting-started/README.md)
- MCP Server for Excel configured and working

## 📁 Files Created

In this tutorial, you'll create:
- `my-first-workbook.xlsx` - Your first automated Excel file

## 🚀 Steps

### Step 1: Create a New Excel File

Let's start by creating a brand new Excel workbook.

**Prompt:**
```
Create a new Excel file at C:\Temp\my-first-workbook.xlsx
```

> 💡 **Tip:** Change the path to match where you want to save files on your computer.

**Expected result:**
- A new Excel file is created
- The AI confirms the file was created successfully

**What's happening:** The MCP Server uses Excel's COM API to create a real Excel workbook. It's not a template or fake file—it's the real thing!

---

### Step 2: Add Some Data

Now let's add data to our workbook. We'll create a simple list.

**Prompt:**
```
Open C:\Temp\my-first-workbook.xlsx and add the following data starting at cell A1:
- Name, Age, City (as headers)
- Alice, 28, New York
- Bob, 35, Los Angeles  
- Carol, 42, Chicago
Then save and close the file.
```

**Expected result:**
The file now contains:

| A | B | C |
|---|---|---|
| Name | Age | City |
| Alice | 28 | New York |
| Bob | 35 | Los Angeles |
| Carol | 42 | Chicago |

---

### Step 3: Read Data Back

Let's verify our data was saved correctly by reading it back.

**Prompt:**
```
Read the contents of C:\Temp\my-first-workbook.xlsx from range A1:C4
```

**Expected result:**
The AI should show you the data in a table or list format, confirming what was written.

---

### Step 4: Modify Existing Data

Now let's update some of the data.

**Prompt:**
```
Open C:\Temp\my-first-workbook.xlsx and change Bob's age to 36, then save and close.
```

**Expected result:**
- The file is opened
- Cell B3 is updated from 35 to 36
- The file is saved and closed

---

### Step 5: Add Formulas

Excel's real power is in formulas. Let's add one!

**Prompt:**
```
Open C:\Temp\my-first-workbook.xlsx and:
1. Add "Average Age" in cell A6
2. Add a formula in B6 that calculates the average of the ages (B2:B4)
3. Save and close
```

**Expected result:**
- Cell A6 contains "Average Age"
- Cell B6 contains `=AVERAGE(B2:B4)` and displays approximately 35.33

---

### Step 6: Batch Operations (Advanced)

For multiple operations, using batch mode is more efficient:

**Prompt:**
```
Using batch mode, open C:\Temp\my-first-workbook.xlsx and:
1. Add a new column D with header "Country" in D1
2. Add "USA" in cells D2, D3, and D4
3. Format the header row (A1:D1) as bold
4. Save and close
```

**Expected result:**
- All operations complete quickly in a single session
- The workbook now has a Country column with formatted headers

---

## ✅ Summary

Congratulations! You've learned the fundamental operations:

| Operation | What You Learned |
|-----------|-----------------|
| **Create** | Make new Excel workbooks from scratch |
| **Write** | Add data to specific cells and ranges |
| **Read** | Retrieve data from Excel files |
| **Update** | Modify existing cell values |
| **Formulas** | Add Excel formulas that calculate automatically |
| **Batch** | Combine multiple operations efficiently |

## 💡 Tips for Better Results

1. **Be specific about file paths** - Use full paths like `C:\Temp\file.xlsx`
2. **Mention save and close** - Always tell the AI to save when you're done
3. **Use batch mode** - For 3+ operations, batch mode is faster
4. **Describe ranges clearly** - Use Excel notation like `A1:C10` or `A1` for single cells

## 🔧 Troubleshooting

### "File not found"
- Make sure the folder exists (e.g., `C:\Temp`)
- Check for typos in the file path

### "File is locked"
- Close Excel if it has the file open
- Check if another process is using the file

### "Data not saved"
- Always include "save" in your prompt before closing
- Verify the save was successful before closing

## ➡️ Next Steps

Ready to learn more? Continue to:

**[Tutorial 3: Working with Ranges →](../03-working-with-ranges/README.md)**

Learn to work with multiple cells, formatting, and more advanced data operations!

---

## 🧪 Practice Exercises

Try these on your own:

1. **Exercise 1:** Create a workbook with your weekly schedule (days as columns, times as rows)

2. **Exercise 2:** Create a simple budget tracker with income, expenses, and a formula for the balance

3. **Exercise 3:** Create a contact list with Name, Phone, and Email columns, then read it back to verify

# Tutorial 1: Installation & Setup

> Get MCP Server for Excel up and running in under 5 minutes

## 🎯 Learning Objectives

By the end of this tutorial, you will be able to:
- Install MCP Server for Excel
- Configure your AI assistant to use it
- Verify everything is working correctly

## ⏱️ Time Required

Approximately 5-10 minutes

## 📋 Prerequisites

- **Windows 10 or 11**
- **Microsoft Excel** installed (Microsoft 365, Excel 2019, or Excel 2016)
- One of the following AI assistants:
  - VS Code with GitHub Copilot
  - Claude Desktop
  - Cursor
  - Windsurf

## 🚀 Installation Options

Choose the option that matches your AI assistant:

---

### Option A: VS Code with GitHub Copilot (Recommended)

The easiest way to get started!

#### Step 1: Install the Extension

1. Open VS Code
2. Go to the Extensions view (`Ctrl+Shift+X`)
3. Search for "Excel MCP Server"
4. Click **Install** on the extension by `sbroenne`

![Install Extension](assets/vscode-install.png)

#### Step 2: Verify Installation

Open GitHub Copilot Chat and try this prompt:

```
List Excel tools available
```

**Expected result:** Copilot should list the Excel MCP tools (excel_file, excel_range, excel_worksheet, etc.)

---

### Option B: Claude Desktop

#### Step 1: Download the MCP Server

1. Go to the [Releases page](https://github.com/sbroenne/mcp-server-excel/releases)
2. Download `ExcelMcp.McpServer-win-x64.zip`
3. Extract to a folder (e.g., `C:\Tools\ExcelMcp`)

#### Step 2: Configure Claude Desktop

1. Open Claude Desktop
2. Go to **Settings** → **Developer** → **Edit Config**
3. Add this configuration:

```json
{
  "mcpServers": {
    "excel": {
      "command": "C:\\Tools\\ExcelMcp\\Sbroenne.ExcelMcp.McpServer.exe"
    }
  }
}
```

4. Restart Claude Desktop

#### Step 3: Verify Installation

Ask Claude:

```
What Excel tools do you have access to?
```

**Expected result:** Claude should describe the available Excel tools.

---

### Option C: Cursor

#### Step 1: Download the MCP Server

Same as Claude Desktop - download and extract from the [Releases page](https://github.com/sbroenne/mcp-server-excel/releases).

#### Step 2: Configure Cursor

1. Open Cursor Settings (`Ctrl+,`)
2. Search for "MCP"
3. Click "Edit in settings.json"
4. Add the configuration:

```json
{
  "mcp": {
    "servers": {
      "excel": {
        "command": "C:\\Tools\\ExcelMcp\\Sbroenne.ExcelMcp.McpServer.exe"
      }
    }
  }
}
```

5. Restart Cursor

---

### Option D: Windsurf

#### Step 1: Download the MCP Server

Same as above - download from the [Releases page](https://github.com/sbroenne/mcp-server-excel/releases).

#### Step 2: Configure Windsurf

1. Open Windsurf
2. Open Command Palette (`Ctrl+Shift+P`)
3. Run "Windsurf: Open MCP Config"
4. Add:

```json
{
  "mcpServers": {
    "excel": {
      "command": "C:\\Tools\\ExcelMcp\\Sbroenne.ExcelMcp.McpServer.exe"
    }
  }
}
```

5. Restart Windsurf

---

## ✅ Verification Test

Once installed, run this test to make sure everything works:

**Prompt:**
```
Create a new Excel file called test.xlsx, add "Hello World" to cell A1, then close it
```

**Expected result:**
1. Excel should briefly open (or run in background)
2. A file called `test.xlsx` should be created
3. Opening it manually should show "Hello World" in cell A1

## 🔧 Troubleshooting

### "Excel tools not available"

- **VS Code**: Make sure the extension is enabled and reload the window
- **Other clients**: Verify the path to the .exe is correct (use double backslashes `\\`)

### "Excel is not installed"

The MCP Server requires a full installation of Microsoft Excel. Excel Online or the mobile app won't work.

### "Access denied" or permission errors

- Run your AI client as Administrator (right-click → Run as administrator)
- Make sure no other process has the Excel file locked

### Excel window doesn't close

This can happen if Excel shows a dialog. Check if Excel is waiting for input in the taskbar.

## ➡️ Next Steps

Now that you're set up, continue to:

**[Tutorial 2: Your First Excel Automation →](../02-first-automation/README.md)**

Learn to create, read, and modify Excel files with natural language!

---

## 📚 Additional Resources

- [Full Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)
- [Configuration Examples](https://github.com/sbroenne/mcp-server-excel/tree/main/examples/mcp-configs)
- [Troubleshooting Guide](https://github.com/sbroenne/mcp-server-excel#troubleshooting)

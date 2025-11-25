# Contributing to MCP Server for Excel Tutorials

Thank you for your interest in contributing! This guide will help you add or improve tutorials.

## 📝 Tutorial Structure

Each tutorial should follow this structure:

```
tutorials/XX-tutorial-name/
├── README.md           # Main tutorial content
├── assets/            # Images, diagrams
│   └── screenshot.png
└── files/             # Sample Excel files
    └── sample.xlsx
```

## ✍️ Writing Guidelines

### Tutorial README Template

```markdown
# Tutorial Title

> Brief one-line description of what you'll learn

## 🎯 Learning Objectives

By the end of this tutorial, you will be able to:
- Objective 1
- Objective 2
- Objective 3

## ⏱️ Time Required

Approximately X minutes

## 📋 Prerequisites

- List any required tutorials to complete first
- Required software or files

## 📁 Files Used

| File | Description |
|------|-------------|
| [sample.xlsx](files/sample.xlsx) | Description of the file |

## 🚀 Steps

### Step 1: Title

Explanation of what we're doing and why.

**Prompt to use:**
```
Your prompt here
```

**Expected result:**
Description or screenshot of what should happen.

### Step 2: Title

Continue with more steps...

## ✅ Summary

What you learned in this tutorial.

## ➡️ Next Steps

- Link to next tutorial
- Additional resources

## 🔧 Troubleshooting

### Common Issue 1
Solution...

### Common Issue 2
Solution...
```

### Style Guidelines

1. **Be conversational** - Write as if you're teaching a friend
2. **Show, don't tell** - Include screenshots and examples
3. **Provide context** - Explain why, not just how
4. **Test everything** - Verify all prompts work with current version
5. **Keep it focused** - One concept per tutorial

### Prompt Guidelines

When writing prompts for users to try:
- Make prompts self-contained (don't rely on previous context)
- Use realistic scenarios
- Include expected output descriptions
- Note any variations that might occur

## 🔄 Pull Request Process

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b tutorial/your-topic`
3. **Write your tutorial** following the structure above
4. **Test all prompts** with the latest MCP Server version
5. **Add to the main README** if it's a new tutorial
6. **Submit a pull request** with a clear description

### PR Checklist

- [ ] Tutorial follows the standard structure
- [ ] All prompts tested and working
- [ ] Sample files included (if applicable)
- [ ] Screenshots/images are clear and helpful
- [ ] Added to main README.md table
- [ ] No spelling/grammar errors

## 📸 Screenshots

- Use PNG format
- Highlight relevant areas
- Keep file sizes reasonable (compress if needed)
- Use consistent styling

## 🏷️ Difficulty Levels

- **Beginner**: No prior Excel automation experience required
- **Intermediate**: Familiarity with Excel and basic prompts
- **Advanced**: Complex scenarios, multiple tools, Power Query/DAX

## ❓ Questions?

Open an issue if you have questions about contributing!

# md2hml - Markdown to HWPML Converter

Convert standard Markdown (`.md`) files into Hancom Office HWPML (`.hml`) format, preserving tables, code blocks, images, and rich text formatting.

This project provides both **Python** and **Node.js** scripts to perform the conversion.

## Features

- **Native HWP Tables**: Converts standard Markdown tables into fully native, editable HWP tables.
- **Code Blocks**: Renders fenced code blocks as boxed tables with monospaced font, mimicking VS Code / GitHub style.
- **Images**: Embeds local images directly into the HWPML file (Base64).
- **Typography**: Supports Bold (`**`), Italic (`*`, `_`), and Inline Code (` ` `).
- **Lists**: Supports nested ordered/unordered lists and task lists (`[ ]`, `[x]`).
- **Blockquotes**: Renders blockquotes with indentation and styling.

## Usage

### Option 1: Python

Requires Python 3.x. No external dependencies required.

```bash
python3 md2hml.py <input_file.md> <output_file.hml>
```

**Example:**
```bash
python3 md2hml.py README.md readme.hml
```

### Option 2: Node.js

Requires Node.js installed. No standard library dependencies required (uses `fs`, `path`).

```bash
node md2hml.js <input_file.md> <output_file.hml>
```

**Example:**
```bash
node md2hml.js README.md readme.hml
```

### Option 3: TypeScript

Requires `typescript` (and optionally `ts-node`).

**Compile and Run:**
```bash
tsc md2hml.ts
node md2hml.js <input_file.md> <output_file.hml>
```

**Run Directly (with ts-node):**
```bash
npx ts-node md2hml.ts <input_file.md> <output_file.hml>
```

## Opening the Output

1.  Run the conversion command.
2.  Open the resulting `.hml` file with **Hancom Office HWP** (HWP 2018 or newer recommended).
3.  You can modify content, resize tables, and save as `.hwp` if needed.

## Markdown Support Reference

| Feature | Markdown | HWP Output |
| :--- | :--- | :--- |
| **Headers** | `# Title` | Heading 1 Style |
| **Bold** | `**text**` | Bold Text |
| **Italic** | `*text*` | Italic Text |
| **Code** | `` `text` `` | Monospaced Text |
| **Lists** | `* Item` | Bullet List |
| **Tables** | `| A | B |` | Native HWP Table |
| **Blockquotes**| `> Text` | Indented Block |
| **Images** | `![alt](path)` | Embedded Image |
| **Horizontal Rule** | `---` | Bottom Border Paragraph |

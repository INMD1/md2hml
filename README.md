
# MD2HML: Markdown to HWPML Converter

MD2HML is a powerful tool that converts Markdown documents into HWPML (Hangul Word Processor Markup Language) format (`.hml`), allowing you to open Markdown content directly in Hancom Office HWP.

This project now includes a **Web Application** and **API** for easy conversion.

> **Note**: The original documentation has been moved to [`original.md`](original.md).

## Features

- **Text Formatting**: Headers, Paragraphs, Bold, Italic.
- **Lists**: Nested lists (up to 6 levels).
- **Links**: Clickable hyperlinks.
- **Code Blocks**: Fenced code blocks with monospace font and background styling.
- **Page Breaks**: Support for `---` as a page break/divider.
- **Images**: Embeds local images (JPEG/PNG) directly into the HML file.
- **Web Interface**: Modern, Drag & Drop UI.
- **API**: REST API for programmatic conversion.

## Usage (CLI Script)

You can run the Python script directly to convert files.

```bash
python3 md2hml.py [input.md] [output.hml]
```

**Example:**
```bash
python3 md2hml.py README.md output.hml
```

## Web Application & API

The project provides a Next.js-based Web Application.

### 1. Installation

Navigate to the `web` directory and install dependencies:

```bash
cd web
npm install
```

### 2. Running the Development Server

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser. You will see the MD2HML Converter interface.

### 3. API Usage

You can use the API to convert Markdown text to HML.

**Endpoint:** `POST /api/convert`

**Body (JSON):**
```json
{
  "markdown": "# Hello World\nThis is a test."
}
```

**Response:** Returns the generated `.hml` file as a download attachment.

## Project Structure

- `md2hml.py`: Core conversion script (Python).
- `web/`: Next.js Web Application source code.
  - `src/app/page.tsx`: Frontend UI.
  - `src/app/api/convert/route.ts`: API Route (wraps `md2hml.py`).
  - `scripts/md2hml.py`: Copy of the script used by the Web App.

## License

MIT

# Frank's Claude Code Skills

A collection of useful skills for Claude Code.

## Skills

| Skill | Description |
|-------|-------------|
| [watermark](./skills/watermark) | Add watermarks to PDF, Word, and Excel files |

## Installation

### Method 1: Clone the repository

```bash
git clone https://github.com/frankfika/franks-claude-code-skills.git
```

### Method 2: Download individual skills

Download the specific skill folder you need from the `skills/` directory.

## Usage

### Watermark Skill

Add watermarks to documents (PDF, Word, Excel).

```bash
# Install dependencies
pip install PyPDF2 reportlab python-docx openpyxl

# Single file
python3 skills/watermark/watermark.py -t "Confidential" document.pdf

# Entire directory
python3 skills/watermark/watermark.py -t "Internal Use Only" -d ./documents

# Output to new directory (preserve originals)
python3 skills/watermark/watermark.py -t "Draft" -d ./docs -o ./watermarked

# Overwrite original files
python3 skills/watermark/watermark.py -t "Confidential" -d ./docs --overwrite
```

**Supported file types:**
- PDF (.pdf)
- Word (.docx)
- Excel (.xlsx)

**Features:**
- Batch processing entire directories
- Preserves directory structure
- Chinese text support
- Multiple watermark positions (center, top-left, bottom-right)

## Contributing

Feel free to submit issues and pull requests!

## License

MIT License

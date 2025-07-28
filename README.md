# DataExtraction

# DocScript

**DocScript** is a Python script that extracts text from DOCX, PPTX, or PDF files (or from direct text input) and generates a segmented video script suitable for narration or voiceover.

## Features

- Extracts text from:
  - Microsoft Word documents (`.docx`)
  - PowerPoint presentations (`.pptx`)
  - PDF files (`.pdf`)
  - Directly pasted text
- Splits extracted text into segments (scenes) for easy video narration
- Saves the generated script to `video_script.txt`

## Requirements

- Python 3.7+
- [python-docx](https://pypi.org/project/python-docx/)
- [python-pptx](https://pypi.org/project/python-pptx/)
- [PyMuPDF](https://pymupdf.readthedocs.io/en/latest/) (`pip install pymupdf`)

Install dependencies with:
```
pip install python-docx python-pptx pymupdf
```

## Usage

1. Run the script:
   ```
   python DocScript.py
   ```

2. Choose your input method:
   - **1**: Upload and extract text from a DOCX, PPTX, or PDF file (enter the full file path)
   - **2**: Paste direct text input (end input with a blank line)

3. The script will generate a segmented video script and save it as `video_script.txt` in the current directory.

## Example

```
Choose input method:
1: Upload and extract text from DOCX/PPTX/PDF
2: Paste direct text input
Enter your choice (1 or 2): 1
Enter full path to your DOCX/PPTX/PDF file: C:\path\to\file.pdf

Generating video script...

Video script saved to video_script.txt

--- Video Script Preview ---

Scene 1:
[First 100 words...]

Scene 2:
[Next 100 words...]
```

## Notes

- Only text is extracted; images are not processed.
- For best results, ensure your input files are not password-protected or corrupted.

---

**Author:** _Aishwarya MS
**License:** Copy Right


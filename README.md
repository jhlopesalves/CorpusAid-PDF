# CorpusAid-PDF

A robust and efficient application for extracting text from PDF files while preserving layout or intelligently handling multi-column documents. Supports various output formats including TXT, HTML, Markdown, and DOCX.

## Features

* **Column-aware Extraction:** Intelligently detects and extracts text from multi-column layouts, maintaining the correct reading order. Ideal for academic papers, magazines, and newspapers.
* **Layout-preserved Extraction:**  Preserves the original document formatting, including spacing, indentation, and special characters. Suitable for forms, technical documents, and code listings.
* **Multiple Output Formats:** Export extracted text as TXT, HTML, Markdown, or DOCX.
* **User-Friendly GUI:**  Intuitive interface for easy file selection, option configuration, and extraction.
* **Batch Processing:** Process multiple PDF files simultaneously.
* **Search Functionality:** Search within the extracted text for specific keywords.
* **PDF Preview:** Preview PDF pages within the application.
* **Zoom and Fit-to-Width:**  Adjust the PDF preview for optimal viewing.
* **Dark/Light Themes:** Choose your preferred theme for a comfortable user experience.
* **Drag and Drop Support:** Drag and drop PDF files directly into the application.
* **Cross-Platform:** Works on Windows, macOS, and Linux.

## Installation

Requires Python 3.7+ and the following libraries:

* `PyMuPDF` (fitz)
* `PySide6`
* `docx`

You can install these dependencies using pip:

```bash
pip install pymupdf PySide6 python-docx
```

## Usage

1. **Open PDF(s):** Use the "Open PDF(s)" button or drag and drop PDF files into the application window.
2. **Select Output Folder:** Choose the destination folder for the extracted text files using the "Save As" button.
3. **Choose Extraction Mode:** Select either "Column-aware" or "Layout-preserved" mode based on the document's structure.
4. **Select Output Format:** Choose the desired output format (TXT, HTML, Markdown, DOCX).
5. **Process PDF(s):** Click the "Process PDF(s)" button to start the extraction process. A progress bar will indicate the progress.
6. **View Results:** The extracted text will be displayed in the "Processed Text" tab. You can also preview the original PDF in the "Original PDF" tab.

## Contributing

Contributions are welcome! Please feel free to submit bug reports, feature requests, or pull requests.

## License

This project is licensed under the [MIT License](LICENSE). (Create a LICENSE file with the MIT license text).

## Screenshots (Optional)

Include screenshots of your application here.

## Contact

For any questions or issues, please contact [jhlopesalves@ufmg.br].

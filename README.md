# 📅 DOCX Date Shifter

A client-side web application that processes Word documents (.docx) to automatically update weekly date ranges. All processing happens in your browser - no files are uploaded to any server, ensuring complete privacy.

## 🌟 Features

- **Client-side Processing**: All file processing happens in your browser - files never leave your computer
- **Batch Processing**: Upload a ZIP file containing multiple .docx files
- **Smart Date Detection**: Automatically detects various date range formats in Word documents
- **Two Update Modes**:
  - **Set New Start Date**: Define a new start date for week 1, and all subsequent weeks are calculated automatically
  - **Shift Dates**: Add or subtract days/weeks/months from all detected dates
- **Preview & Edit**: Review detected date ranges and manually edit them before processing
- **Multiple Date Formats**: Supports numeric dates (DD/MM, DD/MM/YYYY) and month names (15 Sep - 21 Sep)
- **Responsive Design**: Works on desktop and mobile devices

## 🚀 Quick Start

### Option 1: Open Directly (Simplest)

1. Download or clone this repository
2. Open `index.html` in your web browser
3. That's it! The app is ready to use.

### Option 2: Local Development Server

If you prefer using a development server:

```bash
# Clone the repository
git clone https://github.com/your-username/docx-date-shifter.git
cd docx-date-shifter

# Install dependencies (optional, for development)
npm install

# Start development server
npm run dev
```

## 📖 How to Use

1. **Upload Files**: Drag and drop a .zip file containing .docx documents, or click to browse
2. **Choose Update Mode**:
   - **Set New Start Date**: Pick a date for week 1 (e.g., 2025-10-13)
   - **Shift Dates**: Enter how many days/weeks/months to add or subtract
3. **Preview**: Review the detected date ranges in each file
4. **Edit (Optional)**: Manually edit any detected date range if needed
5. **Process**: Click "Process & Download ZIP" to get your updated files

## 🎯 Supported Date Formats

The app automatically detects these date range patterns:

### Numeric Formats

- `15/9-21/9`
- `15/09-21/09`
- `15/9/2025-21/9/2025`
- `15-09-2025 - 21-09-2025`
- `15.09 - 21.09`

### Month Name Formats

- `15 Sep - 21 Sep`
- `1 Oct - 7 Oct`
- `15 Sep 2025 - 21 Sep 2025`

### Arabic Support

The app also supports Arabic month names and range separators (إلى).

## 🔧 Technical Details

### Architecture

- **Frontend Only**: Pure HTML, CSS, and JavaScript
- **No Server Required**: Everything runs in the browser
- **Libraries Used**:
  - [JSZip](https://stuk.github.io/jszip/) - ZIP file handling
  - [FileSaver.js](https://github.com/eligrey/FileSaver.js/) - File downloads
  - [Day.js](https://day.js.org/) - Date manipulation

### How It Works

1. **ZIP Processing**: Extracts .docx files from uploaded ZIP
2. **DOCX Parsing**: Opens each .docx as a ZIP and parses `word/document.xml`
3. **Date Detection**: Uses regex patterns to find date ranges in text content
4. **XML Manipulation**: Updates the Word document's XML structure
5. **Reassembly**: Creates new .docx files and packages them into a downloadable ZIP

### File Size Limits

- Maximum upload size: 50MB
- Supports 40-50+ .docx files in a single ZIP

## 🧪 Testing

### Manual Testing Checklist

1. **Basic Upload**:

   - [ ] Upload a ZIP with .docx files
   - [ ] Verify files are listed and date ranges detected

2. **Set Start Date Mode**:

   - [ ] Set new start date to 2025-10-13
   - [ ] Verify week 1 becomes 13/10-19/10, week 2 becomes 20/10-26/10

3. **Shift Mode**:

   - [ ] Shift by +7 days
   - [ ] Verify all dates move forward by one week

4. **Manual Editing**:

   - [ ] Edit a date range in the preview
   - [ ] Verify the custom change is applied in the output

5. **Error Handling**:
   - [ ] Upload non-ZIP file (should show error)
   - [ ] Upload ZIP without .docx files (should show error)

### Test Files

Use the files in `test-samples/` to test the application:

- `sample-input.zip` - Contains sample .docx files with weekly date ranges

## 🌐 Publishing to GitHub Pages

1. **Create Repository**:

   ```bash
   git init
   git add .
   git commit -m "feat: client-side docx date shifter + demo"
   git branch -M main
   git remote add origin https://github.com/your-username/docx-date-shifter.git
   git push -u origin main
   ```

2. **Enable GitHub Pages**:

   - Go to your repository settings
   - Scroll to "Pages" section
   - Source: Deploy from a branch
   - Branch: `main` / `root`
   - Click "Save"

3. **Access Your App**:
   - Your app will be available at: `https://your-username.github.io/docx-date-shifter/`

## ⚠️ Limitations

- **DOCX Only**: Only supports .docx files (not .doc)
- **Text Formatting**: May alter some text formatting when replacing dates
- **Complex Layouts**: Works best with simple tables and paragraphs
- **Date Assumptions**: Assumes date ranges represent weekly periods

## 🐛 Troubleshooting

### Common Issues

**"No date ranges detected"**

- Check that your document contains date ranges in supported formats
- Ensure dates are in table cells or paragraphs (not in headers/footers)

**"Error processing file"**

- The .docx file may be corrupted or use an unsupported format
- Try opening and re-saving the file in Microsoft Word

**"File too large"**

- Reduce the ZIP file size to under 50MB
- Split large batches into smaller uploads

## 📝 License

MIT License - see [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📞 Support

If you encounter issues:

1. Check the troubleshooting section above
2. Ensure your .docx files contain recognizable date patterns
3. Try with the provided test samples first

---

**Privacy Notice**: This application processes all files locally in your browser. No data is sent to any server or third party.

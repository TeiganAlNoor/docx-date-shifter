// DOCX Date Shifter - Main JavaScript
// Client-side processing of Word documents to update date ranges

// Initialize dayjs with custom parse format plugin
dayjs.extend(dayjs_plugin_customParseFormat);

// Global state
let uploadedZip = null;
let processedFiles = [];
let detectedRanges = new Map(); // filename -> array of range objects

// DOM elements
console.log('🚀 main.js loading...');

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const modeSection = document.getElementById('modeSection');
const previewSection = document.getElementById('previewSection');
const processSection = document.getElementById('processSection');
const resultsSection = document.getElementById('resultsSection');
const fileList = document.getElementById('fileList');
const processBtn = document.getElementById('processBtn');
const progressContainer = document.getElementById('progressContainer');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultsContent = document.getElementById('resultsContent');
const sampleDownload = document.getElementById('sampleDownload');

// Verify all elements are found
console.log('Element check:');
console.log('uploadArea:', uploadArea);
console.log('fileInput:', fileInput);
console.log('processBtn:', processBtn);
console.log('All elements found:', !!(uploadArea && fileInput && processBtn));

// Global variables
const DATE_PATTERNS = {
  // Numeric patterns: DD/MM, D/M, DD/MM/YYYY, etc.
  numeric: [
    /(\d{1,2})[\/\-\.](\d{1,2})(?:[\/\-\.](\d{2,4}))?\s*(?:-|–|—|to|إلى)\s*(\d{1,2})[\/\-\.](\d{1,2})(?:[\/\-\.](\d{2,4}))?/gi,
    // Handle cases with spaces around separators
    /(\d{1,2})\s*[\/\-\.]\s*(\d{1,2})(?:\s*[\/\-\.]\s*(\d{2,4}))?\s*(?:-|–|—|to|إلى)\s*(\d{1,2})\s*[\/\-\.]\s*(\d{1,2})(?:\s*[\/\-\.]\s*(\d{2,4}))?/gi
  ],
  // Month name patterns: 15 Sep - 21 Sep, 1 Oct - 7 Oct
  monthNames: [
    /(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|يناير|فبراير|مارس|أبريل|مايو|يونيو|يوليو|أغسطس|سبتمبر|أكتوبر|نوفمبر|ديسمبر)\s*(\d{2,4})?\s*(?:-|–|—|to|إلى)\s*(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|يناير|فبراير|مارس|أبريل|مايو|يونيو|يوليو|أغسطس|سبتمبر|أكتوبر|نوفمبر|ديسمبر)\s*(\d{2,4})?/gi
  ]
};

// Month name mappings
const MONTH_NAMES = {
  'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
  'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
  'يناير': 1, 'فبراير': 2, 'مارس': 3, 'أبريل': 4, 'مايو': 5, 'يونيو': 6,
  'يوليو': 7, 'أغسطس': 8, 'سبتمبر': 9, 'أكتوبر': 10, 'نوفمبر': 11, 'ديسمبر': 12
};

// Event listeners
document.addEventListener('DOMContentLoaded', () => {
  console.log('🎯 DOM Content Loaded');
  initializeApp();

  // Test that button is clickable
  setTimeout(() => {
    console.log('🧪 Testing button state after 1 second');
    console.log('Button disabled?', processBtn.disabled);
    console.log('Button style display:', processBtn.style.display);
    console.log('Process section display:', processSection.style.display);
  }, 1000);
});

function initializeApp() {
  setupEventListeners();
  setupModeHandlers();
}

function setupEventListeners() {
  // Upload area drag and drop
  uploadArea.addEventListener('click', () => fileInput.click());
  uploadArea.addEventListener('dragover', handleDragOver);
  uploadArea.addEventListener('dragleave', handleDragLeave);
  uploadArea.addEventListener('drop', handleDrop);

  // File input change
  fileInput.addEventListener('change', handleFileSelect);

  // Process button
  console.log('🔧 Setting up process button event listener');
  console.log('processBtn element:', processBtn);
  processBtn.addEventListener('click', () => {
    console.log('🎯 Process button clicked!');
    processFiles();
  });

  // Sample download
  sampleDownload.addEventListener('click', downloadSample);
}

function setupModeHandlers() {
  const modeInputs = document.querySelectorAll('input[name="mode"]');
  modeInputs.forEach(input => {
    input.addEventListener('change', updateModeDisplay);
  });

  // Set default start date to today
  const startDateInput = document.getElementById('startDate');
  startDateInput.value = dayjs().format('YYYY-MM-DD');
}

function updateModeDisplay() {
  const selectedMode = document.querySelector('input[name="mode"]:checked').value;
  const startDateInput = document.getElementById('startDate');
  const shiftControls = document.querySelector('.shift-controls-combined');

  if (selectedMode === 'setStart') {
    startDateInput.style.display = 'block';
    if (shiftControls) shiftControls.style.display = 'none';
  } else {
    startDateInput.style.display = 'none';
    if (shiftControls) shiftControls.style.display = 'grid';
  }
}

// Drag and drop handlers
function handleDragOver(e) {
  e.preventDefault();
  uploadArea.classList.add('dragover');
}

function handleDragLeave(e) {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
}

function handleDrop(e) {
  e.preventDefault();
  uploadArea.classList.remove('dragover');

  const files = e.dataTransfer.files;
  if (files.length > 0) {
    handleFileUpload(files[0]);
  }
}

function handleFileSelect(e) {
  const file = e.target.files[0];
  if (file) {
    handleFileUpload(file);
  }
}

async function handleFileUpload(file) {
  // Validate file
  if (!file.name.toLowerCase().endsWith('.zip')) {
    showError('Please upload a .zip file containing .docx documents.');
    return;
  }

  // No file size limits - allow any size upload
  console.log(`Uploading file: ${file.name}, Size: ${(file.size / 1024 / 1024).toFixed(2)} MB`);

  try {
    showProgress('Loading ZIP file...', 10);

    // Load ZIP file
    uploadedZip = await JSZip.loadAsync(file);

    showProgress('Analyzing documents...', 30);

    // Find all .docx files
    const docxFiles = [];
    uploadedZip.forEach((relativePath, zipEntry) => {
      if (relativePath.toLowerCase().endsWith('.docx') && !relativePath.startsWith('__MACOSX/')) {
        docxFiles.push({
          path: relativePath,
          entry: zipEntry
        });
      }
    });

    if (docxFiles.length === 0) {
      showError('No .docx files found in the uploaded ZIP.');
      return;
    }

    showProgress('Detecting date ranges...', 50);

    // Process each docx file to detect date ranges
    processedFiles = [];
    detectedRanges.clear();

    for (let i = 0; i < docxFiles.length; i++) {
      const file = docxFiles[i];
      showProgress(`Processing ${file.path}...`, 50 + (i / docxFiles.length) * 40);

      try {
        const ranges = await detectDateRangesInDocx(file.entry);
        detectedRanges.set(file.path, ranges);
        processedFiles.push({
          path: file.path,
          entry: file.entry,
          status: ranges.length > 0 ? 'success' : 'warning',
          message: ranges.length > 0 ? `${ranges.length} date ranges found` : 'No date ranges detected'
        });
      } catch (error) {
        console.error(`Error processing ${file.path}:`, error);
        processedFiles.push({
          path: file.path,
          entry: file.entry,
          status: 'error',
          message: `Error: ${error.message}`
        });
        detectedRanges.set(file.path, []);
      }
    }

    hideProgress();
    showPreview();

  } catch (error) {
    console.error('Error processing ZIP file:', error);
    showError(`Error processing ZIP file: ${error.message}`);
    hideProgress();
  }
}

async function detectDateRangesInDocx(zipEntry) {
  try {
    // Extract the docx as a zip
    const docxZip = await JSZip.loadAsync(await zipEntry.async('arraybuffer'));

    // Get document.xml
    const documentXml = await docxZip.file('word/document.xml').async('string');

    // Parse XML
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');

    // Check for parsing errors
    const parserError = xmlDoc.querySelector('parsererror');
    if (parserError) {
      throw new Error('Failed to parse document XML');
    }

    const ranges = [];
    const seenTexts = new Set(); // To avoid duplicates

    // Find all text-containing elements (paragraphs and table cells)
    const textContainers = [
      ...xmlDoc.querySelectorAll('w\\:p, p'), // paragraphs
      ...xmlDoc.querySelectorAll('w\\:tc, tc') // table cells
    ];

    textContainers.forEach((container, containerIndex) => {
      // Get all text runs in this container
      const textRuns = container.querySelectorAll('w\\:t, t');

      // Concatenate all text from runs to get full text
      let fullText = '';
      const runInfo = [];

      textRuns.forEach(run => {
        const text = run.textContent || '';
        runInfo.push({
          element: run,
          startIndex: fullText.length,
          endIndex: fullText.length + text.length,
          text: text
        });
        fullText += text;
      });

      if (fullText.trim()) {
        // Search for date patterns in the full text
        const foundRanges = findDatePatternsInText(fullText);

        foundRanges.forEach(range => {
          // Create a unique key to avoid duplicates
          const uniqueKey = `${range.originalText}-${containerIndex}-${range.startIndex}`;

          if (!seenTexts.has(uniqueKey)) {
            seenTexts.add(uniqueKey);

            ranges.push({
              ...range,
              containerIndex,
              container,
              runInfo,
              fullText,
              uniqueKey,
              // Store the actual text indices for replacement
              textStartIndex: range.startIndex,
              textEndIndex: range.endIndex
            });
          }
        });
      }
    });

    // Sort ranges by their position in the document
    ranges.sort((a, b) => {
      if (a.containerIndex !== b.containerIndex) {
        return a.containerIndex - b.containerIndex;
      }
      return a.startIndex - b.startIndex;
    });

    console.log(`Detected ${ranges.length} unique date ranges`);
    return ranges;

  } catch (error) {
    console.error('Error detecting date ranges:', error);
    throw error;
  }
}

function findDatePatternsInText(text) {
  const ranges = [];

  // Try numeric patterns
  DATE_PATTERNS.numeric.forEach(pattern => {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      const parsed = parseNumericDateRange(match);
      if (parsed) {
        ranges.push({
          originalText: match[0],
          startIndex: match.index,
          endIndex: match.index + match[0].length,
          startDate: parsed.startDate,
          endDate: parsed.endDate,
          pattern: 'numeric'
        });
      }
    }
  });

  // Try month name patterns
  DATE_PATTERNS.monthNames.forEach(pattern => {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      const parsed = parseMonthNameDateRange(match);
      if (parsed) {
        ranges.push({
          originalText: match[0],
          startIndex: match.index,
          endIndex: match.index + match[0].length,
          startDate: parsed.startDate,
          endDate: parsed.endDate,
          pattern: 'monthName'
        });
      }
    }
  });

  return ranges;
}

function parseNumericDateRange(match) {
  try {
    // Extract components: [full, d1, m1, y1, d2, m2, y2]
    const [, d1, m1, y1, d2, m2, y2] = match;

    const currentYear = new Date().getFullYear();

    // Parse start date
    let startDay = parseInt(d1);
    let startMonth = parseInt(m1);
    let startYear = y1 ? (y1.length === 2 ? 2000 + parseInt(y1) : parseInt(y1)) : currentYear;

    // Parse end date
    let endDay = parseInt(d2);
    let endMonth = parseInt(m2);
    let endYear = y2 ? (y2.length === 2 ? 2000 + parseInt(y2) : parseInt(y2)) : currentYear;

    // Validate dates
    if (startDay < 1 || startDay > 31 || startMonth < 1 || startMonth > 12 ||
      endDay < 1 || endDay > 31 || endMonth < 1 || endMonth > 12) {
      return null;
    }

    const startDate = dayjs(`${startYear}-${startMonth.toString().padStart(2, '0')}-${startDay.toString().padStart(2, '0')}`);
    const endDate = dayjs(`${endYear}-${endMonth.toString().padStart(2, '0')}-${endDay.toString().padStart(2, '0')}`);

    if (!startDate.isValid() || !endDate.isValid()) {
      return null;
    }

    return { startDate, endDate };
  } catch (error) {
    return null;
  }
}

function parseMonthNameDateRange(match) {
  try {
    // Extract components: [full, d1, month1, y1, d2, month2, y2]
    const [, d1, month1, y1, d2, month2, y2] = match;

    const currentYear = new Date().getFullYear();

    // Parse start date
    const startDay = parseInt(d1);
    const startMonth = MONTH_NAMES[month1.toLowerCase()];
    const startYear = y1 ? parseInt(y1) : currentYear;

    // Parse end date
    const endDay = parseInt(d2);
    const endMonth = MONTH_NAMES[month2.toLowerCase()];
    const endYear = y2 ? parseInt(y2) : currentYear;

    if (!startMonth || !endMonth) {
      return null;
    }

    const startDate = dayjs(`${startYear}-${startMonth.toString().padStart(2, '0')}-${startDay.toString().padStart(2, '0')}`);
    const endDate = dayjs(`${endYear}-${endMonth.toString().padStart(2, '0')}-${endDay.toString().padStart(2, '0')}`);

    if (!startDate.isValid() || !endDate.isValid()) {
      return null;
    }

    return { startDate, endDate };
  } catch (error) {
    return null;
  }
}

function showPreview() {
  // Clear previous content
  fileList.innerHTML = '';

  // Show relevant sections
  modeSection.style.display = 'block';
  previewSection.style.display = 'block';
  processSection.style.display = 'block';

  // Update mode display
  updateModeDisplay();

  processedFiles.forEach(file => {
    const fileDiv = document.createElement('div');
    fileDiv.className = 'file-item fade-in';

    const ranges = detectedRanges.get(file.path) || [];

    // Count unique date patterns
    const uniquePatterns = new Set(ranges.map(r => r.originalText));
    const duplicateInfo = ranges.length > uniquePatterns.size ?
      ` (${ranges.length} total detections, ${ranges.length - uniquePatterns.size} duplicates removed)` : '';

    fileDiv.innerHTML = `
            <div class="file-header">
                <div class="file-name">${file.path}</div>
                <div class="file-status status-${file.status}">
                    ${uniquePatterns.size} unique weeks found${duplicateInfo}
                </div>
            </div>
            ${ranges.length > 0 ? `
                <div class="date-ranges">
                    <table>
                        <thead>
                            <tr>
                                <th>Original Text</th>
                                <th>Start Date</th>
                                <th>End Date</th>
                                <th>Will Change To</th>
                                <th>Manual Edit</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${ranges.map((range, index) => `
                                <tr>
                                    <td><code>${range.originalText}</code></td>
                                    <td>${range.startDate.format('DD/MM/YYYY')}</td>
                                    <td>${range.endDate.format('DD/MM/YYYY')}</td>
                                    <td><strong style="color: #667eea;">${range.replacement || range.originalText}</strong></td>
                                    <td>
                                        <input type="text" 
                                               class="replacement-input" 
                                               data-file="${file.path}" 
                                               data-range="${index}" 
                                               value="${range.replacement || range.originalText}"
                                               placeholder="Custom replacement...">
                                    </td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            ` : '<p>No date ranges detected in this file.</p>'}
        `;

    fileList.appendChild(fileDiv);
  });

  // Calculate initial replacements based on current mode
  updateReplacementPreviews();

  // Add event listeners for manual edits
  document.querySelectorAll('.replacement-input').forEach(input => {
    input.addEventListener('input', (e) => {
      // Update the replacement in our data
      const filePath = e.target.dataset.file;
      const rangeIndex = parseInt(e.target.dataset.range);
      const ranges = detectedRanges.get(filePath);
      if (ranges && ranges[rangeIndex]) {
        ranges[rangeIndex].replacement = e.target.value;
      }
    });
  });
}

function updateReplacementPreviews() {
  const mode = document.querySelector('input[name="mode"]:checked').value;

  detectedRanges.forEach((ranges, filePath) => {
    // Group ranges by their original text to handle duplicates
    const uniqueRanges = new Map();

    ranges.forEach(range => {
      if (!uniqueRanges.has(range.originalText)) {
        uniqueRanges.set(range.originalText, []);
      }
      uniqueRanges.get(range.originalText).push(range);
    });

    console.log(`File: ${filePath}, Unique date patterns: ${uniqueRanges.size}`);

    // Sort unique ranges by their original dates to get proper week order
    const sortedUniqueTexts = Array.from(uniqueRanges.keys()).sort((a, b) => {
      // Find the first range for each text to get the date
      const rangeA = uniqueRanges.get(a)[0];
      const rangeB = uniqueRanges.get(b)[0];

      if (rangeA.startDate && rangeB.startDate) {
        return rangeA.startDate.isBefore(rangeB.startDate) ? -1 : 1;
      }
      return a.localeCompare(b);
    });

    // Calculate replacements for each unique pattern
    console.log(`\n=== Week Assignment Debug ===`);
    console.log(`File: ${filePath}`);
    console.log(`Total unique patterns: ${sortedUniqueTexts.length}`);

    sortedUniqueTexts.forEach((originalText, weekIndex) => {
      const rangesForThisText = uniqueRanges.get(originalText);
      const firstRange = rangesForThisText[0];

      console.log(`\nWeek ${weekIndex}: "${originalText}"`);
      console.log(`  Start date: ${firstRange.startDate ? firstRange.startDate.format('DD/MM/YYYY') : 'undefined'}`);
      console.log(`  End date: ${firstRange.endDate ? firstRange.endDate.format('DD/MM/YYYY') : 'undefined'}`);

      let replacement = originalText;

      if (mode === 'setStart') {
        const startDate = document.getElementById('startDate').value;
        if (startDate) {
          replacement = calculateSetStartReplacement(firstRange, startDate, weekIndex);
        }
      } else {
        // Get values from all three shift inputs
        const months = parseInt(document.getElementById('shiftMonths').value) || 0;
        const weeks = parseInt(document.getElementById('shiftWeeks').value) || 0;
        const days = parseInt(document.getElementById('shiftDays').value) || 0;

        if (months !== 0 || weeks !== 0 || days !== 0) {
          replacement = calculateCombinedShiftReplacement(firstRange, months, weeks, days);
        }
      }

      // Apply the same replacement to all instances of this text
      rangesForThisText.forEach((range, instanceIndex) => {
        range.replacement = replacement;

        // Update input field
        const input = document.querySelector(`input[data-file="${filePath}"][data-range="${ranges.indexOf(range)}"]`);
        if (input) {
          input.value = replacement;
        }

        // Update the preview column
        const previewCell = document.querySelector(`tr:has(input[data-file="${filePath}"][data-range="${ranges.indexOf(range)}"]) td:nth-child(4)`);
        if (previewCell) {
          previewCell.innerHTML = `<strong style="color: #667eea;">${replacement}</strong>`;
        }
      });

      console.log(`Week ${weekIndex + 1}: "${originalText}" -> "${replacement}" (${rangesForThisText.length} instances)`);
    });
  });

  // Update the changes summary
  updateChangesSummary();
}

function updateChangesSummary() {
  const summaryContainer = document.getElementById('processSummary');
  const summaryContent = document.getElementById('changesSummary');

  if (!summaryContainer || !summaryContent) return;

  let totalChanges = 0;
  let summaryText = '';

  detectedRanges.forEach((ranges, filePath) => {
    const changedRanges = ranges.filter(r => r.replacement && r.replacement !== r.originalText);
    if (changedRanges.length > 0) {
      summaryText += `📄 ${filePath}:\n`;
      changedRanges.forEach(range => {
        summaryText += `  • "${range.originalText}" → "${range.replacement}"\n`;
        totalChanges++;
      });
      summaryText += '\n';
    }
  });

  if (totalChanges === 0) {
    summaryText = 'No changes will be made. All date ranges will remain the same.';
    summaryContainer.style.display = 'none';
  } else {
    summaryText = `Total changes: ${totalChanges}\n\n${summaryText}`;
    summaryContainer.style.display = 'block';
  }

  summaryContent.textContent = summaryText;
}

function calculateSetStartReplacement(range, newStartDateStr, weekNumber) {
  try {
    console.log(`\n--- Week Calculation Debug ---`);
    console.log(`Original range: "${range.originalText}"`);
    console.log(`Original start: ${range.startDate ? range.startDate.format('DD/MM/YYYY') : 'undefined'}`);
    console.log(`Original end: ${range.endDate ? range.endDate.format('DD/MM/YYYY') : 'undefined'}`);
    console.log(`New start date string: "${newStartDateStr}"`);
    console.log(`Week number: ${weekNumber}`);

    const newStartDate = dayjs(newStartDateStr);
    console.log(`Parsed new start date: ${newStartDate.format('DD/MM/YYYY')}`);

    // Calculate the start date for this specific week
    const thisWeekStartDate = newStartDate.add(weekNumber * 7, 'day');
    console.log(`This week start date: ${thisWeekStartDate.format('DD/MM/YYYY')} (added ${weekNumber * 7} days)`);

    // Calculate the duration of the original range
    const originalDuration = range.endDate.diff(range.startDate, 'day');
    console.log(`Original duration: ${originalDuration} days`);

    // Calculate the end date maintaining the same duration
    const thisWeekEndDate = thisWeekStartDate.add(originalDuration, 'day');
    console.log(`This week end date: ${thisWeekEndDate.format('DD/MM/YYYY')}`);

    const result = formatDateRange(thisWeekStartDate, thisWeekEndDate, range.pattern, range.originalText);
    console.log(`Final formatted result: "${result}"`);
    console.log(`--- End Debug ---\n`);

    return result;
  } catch (error) {
    console.error('Error calculating set start replacement:', error);
    return range.originalText;
  }
}

function calculateShiftReplacement(range, amount, unit) {
  try {
    const newStartDate = range.startDate.add(amount, unit);
    const newEndDate = range.endDate.add(amount, unit);

    return formatDateRange(newStartDate, newEndDate, range.pattern, range.originalText);
  } catch (error) {
    return range.originalText;
  }
}

function formatDateRange(startDate, endDate, pattern, originalText) {
  console.log(`\n=== Format Debug ===`);
  console.log(`Start date object: ${startDate.format('DD/MM/YYYY')}`);
  console.log(`End date object: ${endDate.format('DD/MM/YYYY')}`);
  console.log(`Pattern: ${pattern}`);
  console.log(`Original text: "${originalText}"`);

  // Analyze the original format more carefully
  const original = originalText.toLowerCase();

  // Detect separator style
  let separator = '-';
  if (originalText.includes(' - ')) {
    separator = ' - ';
  } else if (originalText.includes('–')) {
    separator = '–';
  } else if (originalText.includes('—')) {
    separator = '—';
  } else if (originalText.includes(' to ')) {
    separator = ' to ';
  } else if (originalText.includes('إلى')) {
    separator = ' إلى ';
  }

  if (pattern === 'numeric') {
    // Detect if original has year
    const hasYear = /\d{4}/.test(originalText);

    // Detect date separator (/, -, .)
    let dateSep = '/';
    if (originalText.includes('-') && !originalText.includes(' - ')) {
      dateSep = '-';
    } else if (originalText.includes('.')) {
      dateSep = '.';
    }

    // Detect if it's DD/MM or D/M format
    const hasLeadingZeros = /\b0\d/.test(originalText);

    if (hasYear) {
      if (hasLeadingZeros) {
        const result = `${startDate.format(`DD${dateSep}MM${dateSep}YYYY`)}${separator}${endDate.format(`DD${dateSep}MM${dateSep}YYYY`)}`;
        console.log(`Format result (hasYear + hasLeadingZeros): "${result}"`);
        console.log(`=== End Format Debug ===\n`);
        return result;
      } else {
        const result = `${startDate.format(`D${dateSep}M${dateSep}YYYY`)}${separator}${endDate.format(`D${dateSep}M${dateSep}YYYY`)}`;
        console.log(`Format result (hasYear, no leading zeros): "${result}"`);
        console.log(`=== End Format Debug ===\n`);
        return result;
      }
    } else {
      if (hasLeadingZeros) {
        const result = `${startDate.format(`DD${dateSep}MM`)}${separator}${endDate.format(`DD${dateSep}MM`)}`;
        console.log(`Format result (no year + hasLeadingZeros): "${result}"`);
        console.log(`=== End Format Debug ===\n`);
        return result;
      } else {
        const result = `${startDate.format(`D${dateSep}M`)}${separator}${endDate.format(`D${dateSep}M`)}`;
        console.log(`Format result (no year, no leading zeros): "${result}"`);
        console.log(`=== End Format Debug ===\n`);
        return result;
      }
    }
  } else if (pattern === 'monthName') {
    // Check if original has year
    const hasYear = /\d{4}/.test(originalText);

    if (hasYear) {
      return `${startDate.format('D MMM YYYY')}${separator}${endDate.format('D MMM YYYY')}`;
    } else {
      return `${startDate.format('D MMM')}${separator}${endDate.format('D MMM')}`;
    }
  }

  // Fallback - try to match the original format as closely as possible
  if (originalText.includes('/')) {
    return `${startDate.format('D/M')}${separator}${endDate.format('D/M')}`;
  } else {
    return `${startDate.format('DD/MM/YYYY')}${separator}${endDate.format('DD/MM/YYYY')}`;
  }
}

// Add event listeners to update previews when mode changes
document.addEventListener('change', (e) => {
  if (e.target.name === 'mode' || e.target.id === 'startDate' ||
    e.target.id === 'shiftMonths' || e.target.id === 'shiftWeeks' || e.target.id === 'shiftDays') {
    updateReplacementPreviews();
  }
});

// Add input event listeners for real-time preview updates
document.addEventListener('input', (e) => {
  if (e.target.id === 'startDate' ||
    e.target.id === 'shiftMonths' || e.target.id === 'shiftWeeks' || e.target.id === 'shiftDays') {
    updateReplacementPreviews();
  }
});

async function processFiles() {
  console.log('🚀 processFiles() called!');
  console.log('uploadedZip:', uploadedZip);
  console.log('processedFiles.length:', processedFiles.length);

  if (!uploadedZip || processedFiles.length === 0) {
    console.log('❌ No files to process');
    showError('No files to process.');
    return;
  }

  // Show confirmation dialog with changes summary
  const changesPreview = generateChangesPreview();
  const confirmed = confirm(`You are about to make the following changes:\n\n${changesPreview}\n\nDo you want to continue?`);

  if (!confirmed) {
    return;
  }

  try {
    processBtn.disabled = true;
    showProgress('Starting processing...', 0);
    resultsSection.style.display = 'none';

    const outputZip = new JSZip();
    const results = [];

    // Debug: Log what we're about to process
    console.log('=== PROCESSING DEBUG INFO ===');
    detectedRanges.forEach((ranges, filePath) => {
      console.log(`File: ${filePath}`);
      ranges.forEach((range, index) => {
        console.log(`  Range ${index}: "${range.originalText}" -> "${range.replacement}"`);
        console.log(`    Start: ${range.startDate?.format('DD/MM/YYYY')}, End: ${range.endDate?.format('DD/MM/YYYY')}`);
      });
    });

    for (let i = 0; i < processedFiles.length; i++) {
      const file = processedFiles[i];
      const progress = (i / processedFiles.length) * 100;
      showProgress(`Processing ${file.path}...`, progress);

      try {
        const ranges = detectedRanges.get(file.path) || [];

        if (ranges.length === 0) {
          // No ranges to update, copy original file
          const originalContent = await file.entry.async('arraybuffer');
          outputZip.file(file.path, originalContent);
          results.push({
            type: 'warning',
            message: `${file.path}: No date ranges found, file unchanged`
          });
        } else {
          // Process the file with replacements
          const updatedContent = await processDocxFile(file.entry, ranges);
          outputZip.file(file.path, updatedContent);

          // Count actual changes
          const changedRanges = ranges.filter(r => r.replacement && r.replacement !== r.originalText);
          results.push({
            type: 'success',
            message: `${file.path}: Updated ${changedRanges.length} of ${ranges.length} date range(s)`
          });
        }
      } catch (error) {
        console.error(`Error processing ${file.path}:`, error);
        // Include original file on error
        const originalContent = await file.entry.async('arraybuffer');
        outputZip.file(file.path, originalContent);
        results.push({
          type: 'error',
          message: `${file.path}: Error - ${error.message}, file unchanged`
        });
      }
    }

    showProgress('Generating download...', 95);

    // Generate the final ZIP
    const content = await outputZip.generateAsync({ type: 'blob' });

    // Download the file
    const timestamp = dayjs().format('YYYY-MM-DD_HH-mm-ss');
    saveAs(content, `updated-documents-${timestamp}.zip`);

    console.log('\n=== FINAL ZIP GENERATED ===');
    console.log('Final ZIP size:', content.size, 'bytes');
    console.log('Download filename:', `updated-documents-${timestamp}.zip`);

    hideProgress();
    showResults(results);

  } catch (error) {
    console.error('Error processing files:', error);
    showError(`Error processing files: ${error.message}`);
    hideProgress();
  } finally {
    processBtn.disabled = false;
  }
}

function generateChangesPreview() {
  let preview = '';
  let totalChanges = 0;

  detectedRanges.forEach((ranges, filePath) => {
    const changedRanges = ranges.filter(r => r.replacement && r.replacement !== r.originalText);
    if (changedRanges.length > 0) {
      preview += `📄 ${filePath}:\n`;
      changedRanges.forEach(range => {
        preview += `  • "${range.originalText}" → "${range.replacement}"\n`;
        totalChanges++;
      });
      preview += '\n';
    }
  });

  if (totalChanges === 0) {
    return 'No changes will be made. All date ranges will remain the same.';
  }

  return `Total changes: ${totalChanges}\n\n${preview}`;
}

async function processDocxFile(zipEntry, ranges) {
  try {
    console.log(`\n=== Processing DOCX with absolute minimal changes ===`);

    // Load the docx as a zip
    const docxZip = await JSZip.loadAsync(await zipEntry.async('arraybuffer'));

    // Get document.xml as raw bytes
    const documentXmlBuffer = await docxZip.file('word/document.xml').async('arraybuffer');
    console.log('Original document.xml size (bytes):', documentXmlBuffer.byteLength);

    // Convert to string using exact same encoding
    const documentXml = new TextDecoder('utf-8', { fatal: false }).decode(documentXmlBuffer);

    // Create a map of exact replacements needed - with validation
    const exactReplacements = new Map();
    ranges.forEach(range => {
      if (range.replacement && range.replacement !== range.originalText) {
        // Validate that this looks like a date range before adding
        const isValidDateRange = /\d{1,2}[\/\-\.]\d{1,2}.*(?:-|–|—|to|إلى).*\d{1,2}[\/\-\.]\d{1,2}/i.test(range.originalText);

        if (isValidDateRange && !exactReplacements.has(range.originalText)) {
          exactReplacements.set(range.originalText, range.replacement);
          console.log(`✓ Added valid date range: "${range.originalText}" → "${range.replacement}"`);
        } else if (!isValidDateRange) {
          console.warn(`⚠ Skipping non-date text: "${range.originalText}"`);
        }
      }
    });

    console.log(`Will make ${exactReplacements.size} exact text replacements`);

    let modifiedXml = documentXml;
    let totalReplacements = 0;

    // Perform each replacement one by one with verification
    for (const [originalText, replacement] of exactReplacements) {
      console.log(`\n--- Exact replacement ---`);
      console.log(`"${originalText}" → "${replacement}"`);

      // Check if the original text exists
      if (!modifiedXml.includes(originalText)) {
        console.warn(`⚠ Text "${originalText}" not found - skipping`);
        continue;
      }

      // Count occurrences
      const occurrences = modifiedXml.split(originalText).length - 1;
      console.log(`Found ${occurrences} occurrence(s)`);

      // Only replace if the replacement text is similar in length to prevent layout changes
      const lengthDifference = Math.abs(originalText.length - replacement.length);
      if (lengthDifference > 20) {
        console.warn(`⚠ Large length difference (${lengthDifference} chars) - might affect layout, skipping`);
        continue;
      }

      // Additional safety: ensure the replacement looks like a valid date range
      const replacementIsValidDate = /\d{1,2}[\/\-\.]\d{1,2}.*(?:-|–|—|to|إلى).*\d{1,2}[\/\-\.]\d{1,2}/i.test(replacement);
      if (!replacementIsValidDate) {
        console.warn(`⚠ Replacement doesn't look like a date range: "${replacement}" - skipping`);
        continue;
      }

      // Perform the replacement using the most basic method
      modifiedXml = modifiedXml.replaceAll(originalText, replacement);
      totalReplacements += occurrences;

      // Verify the replacement was successful
      const newOccurrences = modifiedXml.split(replacement).length - 1;
      console.log(`✓ Verified: ${newOccurrences} occurrence(s) of replacement text`);
    }

    console.log(`\nTotal replacements made: ${totalReplacements}`);

    // Convert back to bytes with identical encoding
    const modifiedBuffer = new TextEncoder().encode(modifiedXml);
    console.log('Modified document.xml size (bytes):', modifiedBuffer.byteLength);
    console.log('Byte size change:', modifiedBuffer.byteLength - documentXmlBuffer.byteLength);

    // Clone the original ZIP structure exactly
    const resultZip = docxZip.clone();

    // Replace ONLY the document.xml file, keeping everything else identical
    resultZip.file('word/document.xml', modifiedBuffer, {
      binary: true,
      compression: 'DEFLATE',  // Use same compression as original
      compressionOptions: {
        level: 6  // Standard compression level
      }
    });

    // Generate with settings that preserve the original structure
    const result = await resultZip.generateAsync({
      type: 'arraybuffer',
      compression: 'DEFLATE',
      compressionOptions: {
        level: 6
      },
      streamFiles: false
    });

    console.log('Final DOCX size:', result.byteLength);

    // Verification
    console.log('\n=== Final Verification ===');
    const verifyZip = await JSZip.loadAsync(result);
    const verifyBuffer = await verifyZip.file('word/document.xml').async('arraybuffer');
    const verifyXml = new TextDecoder('utf-8', { fatal: false }).decode(verifyBuffer);

    let verifiedReplacements = 0;
    for (const [originalText, replacement] of exactReplacements) {
      const count = verifyXml.split(replacement).length - 1;
      if (count > 0) {
        verifiedReplacements += count;
        console.log(`✓ "${replacement}": ${count} instances`);
      }
    }

    console.log(`Final verification: ${verifiedReplacements} total replacements confirmed`);

    return result;

  } catch (error) {
    console.error('Error in minimal DOCX processing:', error);
    throw error;
  }
}

// Helper function to escape special regex characters
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Helper function to escape special regex characters
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function showProgress(text, percent) {
  progressContainer.style.display = 'block';
  progressText.textContent = text;
  progressFill.style.width = `${percent}%`;
}

function hideProgress() {
  progressContainer.style.display = 'none';
}

function showResults(results) {
  resultsContent.innerHTML = results.map(result => `
        <div class="result-item result-${result.type}">
            <span>${result.type === 'success' ? '✅' : result.type === 'warning' ? '⚠️' : '❌'}</span>
            <span>${result.message}</span>
        </div>
    `).join('');

  resultsSection.style.display = 'block';
  resultsSection.scrollIntoView({ behavior: 'smooth' });
}

function showError(message) {
  alert(message); // Simple error display - could be enhanced with a modal
}

async function downloadSample() {
  try {
    // Create sample DOCX files
    const sampleZip = new JSZip();

    // Create first sample document
    const docx1 = new JSZip();

    // Add required DOCX structure files
    docx1.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

    docx1.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

    docx1.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`);

    docx1.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>جدول الأسابيع - Sample Weekly Schedule</w:t>
            </w:r>
        </w:p>
        <w:tbl>
            <w:tblPr>
                <w:tblStyle w:val="TableGrid"/>
                <w:tblW w:w="0" w:type="auto"/>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>الأسبوع الأول</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>6/9-12/9</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>سورة المرسلات</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>الأسبوع الثاني</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>13/9-19/9</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>سورة المجادلة الآية (8-1)</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>الأسبوع الثالث</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>20/9-26/9</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>سورة الحشر (21-10)</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
            <w:tr>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>الأسبوع الرابع</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>27/9-3/10</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:p>
                        <w:r>
                            <w:t>سورة الممتحنة (13-6)</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
        <w:p>
            <w:r>
                <w:t>Additional test ranges: 15/9-21/9 and 22/9-28/9</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>`);

    // Generate the first DOCX
    const docx1Blob = await docx1.generateAsync({ type: 'blob' });
    sampleZip.file('sample-schedule.docx', docx1Blob);

    // Create second sample document with different format
    const docx2 = new JSZip();

    // Add the same structure files
    docx2.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

    docx2.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

    docx2.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`);

    docx2.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>Test Document 2 - Monthly Schedule</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Week 1: 15/09/2025 - 21/09/2025</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Week 2: 22/09/2025 - 28/09/2025</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Week 3: 29 Sep - 5 Oct</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Week 4: 6 Oct - 12 Oct</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>`);

    // Generate the second DOCX
    const docx2Blob = await docx2.generateAsync({ type: 'blob' });
    sampleZip.file('test-document-2.docx', docx2Blob);

    // Generate and download the sample ZIP
    const sampleZipBlob = await sampleZip.generateAsync({ type: 'blob' });
    saveAs(sampleZipBlob, 'sample-input.zip');

  } catch (error) {
    console.error('Error creating sample:', error);
    showError('Error creating sample files. Please try again.');
  }
}

/**
 * Calculate and format replacement for combined shift (months + weeks + days)
 * @param {object} range The date range object with startDate and endDate
 * @param {number} months Number of months to shift
 * @param {number} weeks Number of weeks to shift  
 * @param {number} days Number of days to shift
 * @returns {string} Replacement text preserving original format
 */
function calculateCombinedShiftReplacement(range, months, weeks, days) {
  // If no shifts are specified, return original
  if (!months && !weeks && !days) {
    return range.originalText;
  }

  console.log(`Combined shift - Months: ${months}, Weeks: ${weeks}, Days: ${days}`);
  console.log(`Original range:`, range);

  if (!range.startDate || !range.endDate) {
    console.warn('Invalid date range for combined shift:', range);
    return range.originalText;
  }

  // Apply combined shifts to both start and end dates
  let newStartDate = range.startDate;
  let newEndDate = range.endDate;

  if (months) {
    newStartDate = newStartDate.add(parseInt(months), 'month');
    newEndDate = newEndDate.add(parseInt(months), 'month');
  }

  if (weeks) {
    newStartDate = newStartDate.add(parseInt(weeks), 'week');
    newEndDate = newEndDate.add(parseInt(weeks), 'week');
  }

  if (days) {
    newStartDate = newStartDate.add(parseInt(days), 'day');
    newEndDate = newEndDate.add(parseInt(days), 'day');
  }

  // Analyze the original format more carefully
  const originalText = range.originalText;
  const original = originalText.toLowerCase();

  // Detect separator style
  let separator = '-';
  if (originalText.includes(' - ')) {
    separator = ' - ';
  } else if (originalText.includes('–')) {
    separator = '–';
  } else if (originalText.includes('—')) {
    separator = '—';
  } else if (originalText.includes(' to ')) {
    separator = ' to ';
  } else if (originalText.includes(' - ')) {
    separator = ' - ';
  }

  // Create replacement based on detected format
  const replacement = `${newStartDate.format('DD/MM/YYYY')}${separator}${newEndDate.format('DD/MM/YYYY')}`;

  console.log(`Combined shift: ${originalText} → ${replacement}`);
  return replacement;
}

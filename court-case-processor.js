// Court Case Data Processing Utility
class CourtCaseProcessor {
    constructor() {
        // Court keywords for identification
        this.courtKeywords = ["PDJ", "ADJ", "2ADJ", "MSD", "CHIEF", "2SD", "3SD", "4SD", "1JD", "2JD", "3JD", "4JD", "RAN", "KUT", "JJB"];

        // ESTA keywords
        this.estaKeywords = ["PBR", "RAN", "KUT"];

        // ESTA Court Mapping
        this.ESTA_COURT_MAP = {
            "PBR": ["PDJ", "ADJ", "2ADJ", "MSD", "CHIEF", "2SD", "3SD", "4SD", "1JD", "2JD", "3JD", "4JD"],
            "RAN": ["RAN"],
            "KUT": ["KUT"]
        };

        // Column mapping for standardization
        this.columnMapping = {
            'Case No.': 'Cases',
            'Petitioner Name VS Respondent Name': 'Party Name',
            'Next Purpose': 'Purpose'
        };

        // CAT1 to SIDE mapping
        this.cat1ToSideMapping = {
            "MACP": "CIVIL", "ATRO": "CRIMINAL", "TADA": "CRIMINAL",
            "CMA DC": "CIVIL", "SC": "CRIMINAL", "CR A": "CRIMINAL",
            "RCA": "CIVIL", "CR RA": "CRIMINAL", "MACEX": "CIVIL",
            "ACB": "CRIMINAL", "MACMA": "CIVIL", "COMM EX": "CIVIL",
            "PCSO": "CRIMINAL", "GLGP": "CRIMINAL", "SMRY": "CIVIL",
            "EXE S": "CIVIL", "ELEC": "CRIMINAL", "MCA": "CIVIL",
            "CRMA S": "CRIMINAL", "NDPS": "CRIMINAL", "REWA": "CIVIL",
            "GPID CC": "CRIMINAL", "CR EN": "CRIMINAL", "TMSUIT": "CIVIL",
            "CC": "CRIMINAL", "SPCS": "CIVIL", "CMA SC": "CIVIL",
            "RCS": "CIVIL", "CRMA J": "CRIMINAL", "CC JUVE": "CRIMINAL",
            "EXE LAR": "CIVIL", "EXE R": "CIVIL", "SMST R": "CIVIL",
            "SMST S": "CIVIL", "LAR": "CIVIL", "COMM CS": "CIVIL", "HMP": "CIVIL"
        };
    }

    // Get court name from filename
    getCourtName(fileName) {
        for (let keyword of this.courtKeywords) {
            if (fileName.includes(keyword)) {
                return keyword;
            }
        }
        return "PDJ";  // Default court name
    }

    // Get ESTA name from filename and court name
    getEstaName(fileName, courtName) {
        for (let keyword of this.estaKeywords) {
            if (fileName.includes(keyword)) {
                return keyword;
            }
        }

        for (let [esta, courts] of Object.entries(this.ESTA_COURT_MAP)) {
            if (courts.includes(courtName)) {
                return esta;
            }
        }

        return "PBR";  // Default ESTA
    }

    // Process Excel file (remaining methods stay the same as in previous version)
    async processExcelFile(file) {
        try {
            const workbook = await this.readExcelFile(file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            let data = XLSX.utils.sheet_to_json(worksheet);

            const processedData = data.map(row => {
                const processedRow = { ...row };

                for (let [oldKey, newKey] of Object.entries(this.columnMapping)) {
                    if (processedRow[oldKey] !== undefined) {
                        processedRow[newKey] = processedRow[oldKey];
                        delete processedRow[oldKey];
                    }
                }

                processedRow['File Name'] = file.name;
                processedRow['Court Name'] = this.getCourtName(file.name);
                processedRow['ESTA'] = this.getEstaName(file.name, processedRow['Court Name']);
                processedRow['UID'] = `${processedRow['ESTA']}/${processedRow['Cases']}`;

                this.enrichCaseDetails(processedRow);

                return processedRow;
            });

            return processedData;
        } catch (error) {
            console.error(`Error processing file ${file.name}:`, error);
            throw error;
        }
    }

    // Enrichment method with same previous logic
    enrichCaseDetails(details) {
        // Extract CAT1
        const caseNo = details['Cases'] || '';
        details['CAT1'] = caseNo.includes('/') ? caseNo.split('/')[0] : 'UNKNOWN';

        // IPC Special check
        const actSection = details['Act Section'] || '';
        details['IPC SPECIAL'] = ['467', '468', '465', '466', '471'].some(code =>
            actSection.includes(code)) ? 'IPC SPECIAL' : '';

        // CAT2 creation
        details['CAT2'] = `${details['CAT1']}/${details['Nature'] || ''}/${details['IPC SPECIAL'] || ''}`;

        // Date processing
        const processDate = (dateStr) => {
            if (!dateStr) return null;
            try {
                const [day, month, year] = dateStr.split('-');
                return new Date(`${year}-${month}-${day}`);
            } catch {
                return new Date();
            }
        };

        const registrationDate = processDate(details['Date of Registration']);
        const decisionDate = processDate(details['Date of Decision']) || new Date();

        // New date formatting
        details['New Date of Registration'] = registrationDate ?
            registrationDate.toISOString().split('T')[0] : '';
        details['New Date of Decision'] = decisionDate.toISOString().split('T')[0];

        // Status determination
        details['STATUS'] = details['Date of Decision'] ? 'DISPOSE' : 'PENDING';

        // Decision month
        details['DECISION MONTH'] = decisionDate ?
            decisionDate.toLocaleString('default', { month: 'short' }).toUpperCase() +
            '-' + decisionDate.getFullYear().toString().slice(-2) : '';

        // Age calculation
        const ageDays = (decisionDate - registrationDate) / (1000 * 60 * 60 * 24);
        const ageYears = ageDays / 365.25;

        // Special handling for same date registration and decision
        if (registrationDate && decisionDate &&
            registrationDate.toISOString().split('T')[0] === decisionDate.toISOString().split('T')[0]) {
            details['AGE2'] = 0;
            details['AGE CATEGORY'] = '0-5YR';
        } else {
            details['AGE2'] = ageYears > 0 ? Number(ageYears.toFixed(2)) : '';

            // Age category
            if (typeof details['AGE2'] === 'number') {
                if (details['AGE2'] < 5) details['AGE CATEGORY'] = '0-5YR';
                else if (details['AGE2'] < 10) details['AGE CATEGORY'] = '5-10YR';
                else if (details['AGE2'] < 20) details['AGE CATEGORY'] = '10-20YR';
                else if (details['AGE2'] < 30) details['AGE CATEGORY'] = '20-30YR';
                else details['AGE CATEGORY'] = '30YR ABOVE';
            } else {
                details['AGE CATEGORY'] = '';
            }
        }

        // Side determination
        details['SIDE'] = this.cat1ToSideMapping[details['CAT1']] || 'UNKNOWN';

        return details;
    }

    // Generate summary for pending cases
    generatePendingSummary(data) {
        const pendingCases = data.filter(row => row.STATUS === 'PENDING');
        return this.generateCaseSummary(pendingCases);
    }

    // Generate summary for disposed cases
    generateDisposedSummary(data) {
        const disposedCases = data.filter(row => row.STATUS === 'DISPOSE');
        return this.generateCaseSummary(disposedCases);
    }

    // Generate case summary in pivot table format
    generateCaseSummary(data, filename, sheetName) {
        // Get unique CAT2 and AGE CATEGORY values
        const uniqueCAT2 = [...new Set(data.map(row => row.CAT2))];
        const ageCategories = ['0-5YR', '5-10YR', '10-20YR', '20-30YR', '30YR ABOVE'];

        // Create pivot table data structure
        const pivotData = uniqueCAT2.map(cat2 => {
            const row = { CAT2: cat2 };

            // Initialize counts for each age category
            ageCategories.forEach(age => {
                row[age] = data.filter(item =>
                    item.CAT2 === cat2 &&
                    item['AGE CATEGORY'] === age
                ).length;
            });

            // Add total count for the CAT2
            row['Total'] = ageCategories.reduce((sum, age) => sum + row[age], 0);

            return row;
        });

        // Sort by total count descending
        const sortedData = pivotData.sort((a, b) => b.Total - a.Total);

        // Calculate column totals
        const totalsRow = { CAT2: 'TOTAL' };
        ageCategories.forEach(age => {
            totalsRow[age] = sortedData.reduce((sum, row) => sum + row[age], 0);
        });
        totalsRow['Total'] = ageCategories.reduce((sum, age) => sum + totalsRow[age], 0);

        // Add totals row to the sorted data
        const summaryData = [...sortedData, totalsRow];

        return summaryData;
    }

    // Generate pending cases summary
    generatePendingSummary(data) {
        const pendingCases = data.filter(row => row.STATUS === 'PENDING');
        return this.generateCaseSummary(pendingCases, 'pending_cases_summary.xlsx', 'Pending Cases Summary');
    }

    // Generate disposed cases summary
    generateDisposedSummary(data) {
        const disposedCases = data.filter(row => row.STATUS === 'DISPOSE');
        return this.generateCaseSummary(disposedCases, 'disposed_cases_summary.xlsx', 'Disposed Cases Summary');
    }

    // Main processing method
    async processFiles(files) {
        try {
            const processedFiles = await Promise.all(
                Array.from(files).map(file => this.processExcelFile(file))
            );

            const allData = processedFiles.flat();
            const uniqueData = this.removeDuplicates(allData);

            return uniqueData.sort((a, b) => {
                const ageA = parseFloat(a.AGE2) || -Infinity;
                const ageB = parseFloat(b.AGE2) || -Infinity;
                return ageB - ageA;
            });
        } catch (error) {
            console.error('Processing error:', error);
            throw error;
        }
    }

    // Remove duplicates method
    removeDuplicates(data) {
        // Sort data to prioritize DISPOSE over PENDING
        const sortedData = data.sort((a, b) => {
            const statusOrder = { 'DISPOSE': 0, 'PENDING': 1 };
            return statusOrder[a.STATUS] - statusOrder[b.STATUS];
        });

        // Remove duplicates keeping first occurrence of each UID
        const uniqueData = [];
        const seenUIDs = new Set();

        for (const item of sortedData) {
            if (!seenUIDs.has(item.UID)) {
                seenUIDs.add(item.UID);
                uniqueData.push(item);
            }
        }

        return uniqueData;
    }

    // Read Excel file
    async readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                resolve(workbook);
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }
}

// UI Event Handling
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const processButton = document.getElementById('processButton');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    const downloadContainer = document.getElementById('downloadContainer');

    const processor = new CourtCaseProcessor();

    processButton.addEventListener('click', async () => {
        const files = fileInput.files;
        if (files.length === 0) {
            alert('Please select Excel files to process');
            return;
        }

        // Reset download container
        downloadContainer.innerHTML = '';
        progressContainer.style.display = 'block';
        downloadContainer.style.display = 'none';
        progressText.textContent = 'Processing files...';
        progressBar.style.width = '0%';

        try {
            // Process files
            const processedData = await processor.processFiles(files);

            // Create a single workbook with multiple sheets
            const combinedWorkbook = XLSX.utils.book_new();

            // Add Processed Data sheet
            const mainWorksheet = XLSX.utils.json_to_sheet(processedData);
            XLSX.utils.book_append_sheet(combinedWorkbook, mainWorksheet, 'Processed Data');

            // Add Pending Cases Summary sheet
            const pendingSummaryData = processor.generatePendingSummary(processedData);
            const pendingSummaryWorksheet = XLSX.utils.json_to_sheet(pendingSummaryData);
            XLSX.utils.book_append_sheet(combinedWorkbook, pendingSummaryWorksheet, 'Pending Cases Summary');

            // Add Disposed Cases Summary sheet
            const disposedSummaryData = processor.generateDisposedSummary(processedData);
            const disposedSummaryWorksheet = XLSX.utils.json_to_sheet(disposedSummaryData);
            XLSX.utils.book_append_sheet(combinedWorkbook, disposedSummaryWorksheet, 'Disposed Cases Summary');

            // Generate single download link
            const wbout = XLSX.write(combinedWorkbook, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([wbout], { type: 'application/octet-stream' });
            const url = URL.createObjectURL(blob);

            // Create download link
            const downloadLink = document.createElement('a');
            downloadLink.href = url;
            downloadLink.download = 'court_case_reports.xlsx';
            downloadLink.textContent = 'Download Combined Reports';
            downloadLink.className = 'download-link';

            // Add link to container
            downloadContainer.appendChild(downloadLink);
            downloadContainer.style.display = 'flex';

            // Update progress
            progressBar.style.width = '100%';
            progressText.textContent = 'Processing Complete';

        } catch (error) {
            console.error('Error processing files:', error);
            progressText.textContent = 'Processing Failed';
            alert('An error occurred while processing files');
        }
    });
});

// Import SheetJS library
document.write('<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.5/xlsx.full.min.js"><\/script>');

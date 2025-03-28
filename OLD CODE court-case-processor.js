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
        // Try to find court keyword in filename
        for (let keyword of this.courtKeywords) {
            if (fileName.includes(keyword)) {
                return keyword;
            }
        }
        // If no court name found, return default
        return "PDJ";
    }

    // Get ESTA name from filename and court name
    getEstaName(fileName, courtName) {
        // Check ESTA keywords in filename first
        for (let keyword of this.estaKeywords) {
            if (fileName.includes(keyword)) {
                return keyword;
            }
        }

        // If not found in filename, check court name mapping
        for (let [esta, courts] of Object.entries(this.ESTA_COURT_MAP)) {
            if (courts.includes(courtName)) {
                return esta;
            }
        }

        // If no ESTA found, return default
        return "PBR";
    }

    // Process Excel file
    async processExcelFile(file) {
        const workbook = await this.readExcelFile(file);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        let data = XLSX.utils.sheet_to_json(worksheet);

        // Process each row
        const processedData = data.map(row => {
            // Clone the original row to preserve all original columns
            const processedRow = {...row};

            // Rename columns
            for (let [oldKey, newKey] of Object.entries(this.columnMapping)) {
                if (processedRow[oldKey] !== undefined) {
                    processedRow[newKey] = processedRow[oldKey];
                    delete processedRow[oldKey];
                }
            }

            // Add metadata
            processedRow['File Name'] = file.name;
            processedRow['Court Name'] = this.getCourtName(file.name);
            processedRow['ESTA'] = this.getEstaName(file.name, processedRow['Court Name']);
            processedRow['UID'] = `${processedRow['ESTA']}/${processedRow['Cases']}`;

            // Process case details
            this.enrichCaseDetails(processedRow);

            return processedRow;
        });

        return processedData;
    }

    // Enrich case details with additional processing
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

    // Read Excel file
    async readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                resolve(workbook);
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    // Remove duplicates based on UID and STATUS
    removeDuplicates(data) {
        // Sort data to prioritize DISPOSE over PENDING
        const sortedData = data.sort((a, b) => {
            const statusOrder = {'DISPOSE': 0, 'PENDING': 1};
            return statusOrder[a.STATUS] - statusOrder[b.STATUS];
        });

        // Remove duplicates keeping first occurrence of each UID
        const uniqueData = [];
        const seenUIDs = new Set();

        for (const row of sortedData) {
            if (!seenUIDs.has(row.UID)) {
                uniqueData.push(row);
                seenUIDs.add(row.UID);
            }
        }

        return uniqueData;
    }

    // Main processing method
    async processFiles(files) {
        try {
            // Process each file
            const processedFiles = await Promise.all(
                Array.from(files).map(file => this.processExcelFile(file))
            );

            // Flatten processed files
            const allData = processedFiles.flat();

            // Remove duplicates
            const uniqueData = this.removeDuplicates(allData);

            // Sort by AGE2 in descending order
            const sortedData = uniqueData.sort((a, b) => {
                const ageA = parseFloat(a.AGE2) || -Infinity;
                const ageB = parseFloat(b.AGE2) || -Infinity;
                return ageB - ageA;
            });

            return sortedData;
        } catch (error) {
            console.error('Processing error:', error);
            throw error;
        }
    }
}

// UI Event Handling remains the same as in the previous script
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const processButton = document.getElementById('processButton');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    const downloadContainer = document.getElementById('downloadContainer');
    const downloadLink = document.getElementById('downloadLink');

    const processor = new CourtCaseProcessor();

    processButton.addEventListener('click', async () => {
        const files = fileInput.files;
        if (files.length === 0) {
            alert('Please select Excel files to process');
            return;
        }

        // Show progress
        progressContainer.style.display = 'block';
        downloadContainer.style.display = 'none';
        progressText.textContent = 'Processing files...';
        progressBar.style.width = '0%';

        try {
            // Process files
            const processedData = await processor.processFiles(files);

            // Create workbook
            const worksheet = XLSX.utils.json_to_sheet(processedData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Processed Data');

            // Generate download link
            const wbout = XLSX.write(workbook, {bookType: 'xlsx', type: 'array'});
            const blob = new Blob([wbout], {type: 'application/octet-stream'});
            const url = URL.createObjectURL(blob);

            // Update UI
            progressBar.style.width = '100%';
            progressText.textContent = 'Processing Complete';
            downloadContainer.style.display = 'block';
            downloadLink.href = url;
            downloadLink.download = 'processed_court_case_data.xlsx';

        } catch (error) {
            console.error('Error processing files:', error);
            progressText.textContent = 'Processing Failed';
            alert('An error occurred while processing files');
        }
    });
});

// Import SheetJS library
document.write('<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.5/xlsx.full.min.js"><\/script>');

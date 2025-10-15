// --- CONFIGURATION ---
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const JOBS_SHEET_NAME = 'Jobs';
const SETTINGS_SHEET_NAME = 'Settings';
const GEMINI_API_ENDPOINT_GENERATE = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-05-20:generateContent?key=';


/**
 * An installable trigger that runs when a user edits the spreadsheet.
 * It handles checkbox actions (trashing/restoring cover letters) and triggers a resort of the job list.
 * NOTE: This must be installed manually as an "On edit" trigger from the Triggers menu.
 * @param {Object} e The event object from a user's edit.
 */
function handleEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  // Exit if the edit was not on a checkbox in the 'Jobs' sheet.
  if (sheet.getName() !== JOBS_SHEET_NAME || range.getRow() <= 1 || range.getColumn() > 2) {
    return;
  }

  const row = range.getRow();
  const coverLetterCell = sheet.getRange(row, 6); // Column F for Cover Letter
  const formula = coverLetterCell.getFormula();
  
  if (!formula) return; // No link to process

  const urlMatch = formula.match(/=HYPERLINK\("([^"]+)"/);
  if (!urlMatch || !urlMatch[1]) return; // Not a valid hyperlink

  const docUrl = urlMatch[1];
  const docIdMatch = docUrl.match(/d\/([a-zA-Z0-9-_]+)/);
  if (!docIdMatch || !docIdMatch[1]) return; // No doc ID found

  const docId = docIdMatch[1];

  // If "Not Applying" checkbox (column 2) is checked/unchecked.
  if (range.getColumn() === 2) {
    try {
      const file = DriveApp.getFileById(docId);
      if (range.getValue() === true) { // Box is checked: Trash the file
        file.setTrashed(true);
        Logger.log(`Trashed cover letter for job in row ${row}`);
      } else { // Box is unchecked: Restore the file
        file.setTrashed(false);
        const settings = getSettings();
        const folder = getFolderFromUrl(settings.folderLink);
        folder.addFile(file); // Ensure it's in the correct folder
        DriveApp.getRootFolder().removeFile(file); // Clean up from root if it was there
        Logger.log(`Restored cover letter for job in row ${row}`);
      }
    } catch (err) {
      Logger.log(`Could not process cover letter for row ${row}. It may have been permanently deleted. Error: ${err.message}`);
    }
  }

  // Use a small delay to prevent issues with rapid edits, then sort the sheet.
  Utilities.sleep(500);
  sortJobsSheet();
}

/**
 * Creates a custom menu in the spreadsheet UI for easy manual execution.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Job Automator')
    .addItem('1. Run Initial Setup (Formatting)', 'runInitialSetup')
    .addSeparator()
    .addItem('2. Find New Jobs Manually', 'findAndProcessNewJobs')
    .addItem('3. Generate Next Cover Letter', 'generateSingleCoverLetter')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Debug')
        .addItem('Reset Setup Lock', 'resetSetupLock'))
    .addToUi();
}

/**
 * A wrapper function to create and format the necessary sheets, set up conditional formatting, and hide helper columns.
 * This is called from the custom menu and is the first step for a new user.
 */
function runInitialSetup() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('SETUP_COMPLETE') === 'true') {
    SpreadsheetApp.getUi().alert('Setup Already Complete', 'The initial setup has already been run. To prevent accidental data loss, this action is locked. If you need to re-run the setup, please use the "Debug > Reset Setup Lock" menu item.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const ss = SPREADSHEET;

  // --- 1. Ensure our primary sheets exist ---
  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  }

  let jobsSheet = ss.getSheetByName(JOBS_SHEET_NAME);
  if (!jobsSheet) {
    jobsSheet = ss.insertSheet(JOBS_SHEET_NAME, 0); // Insert as the first sheet
  }
  
  // --- 2. Delete any other sheets ---
  const allSheets = ss.getSheets();
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName !== JOBS_SHEET_NAME && sheetName !== SETTINGS_SHEET_NAME) {
      ss.deleteSheet(sheet);
    }
  });


  // --- 3. Set up the Settings Sheet ---
  settingsSheet.clear(); // Clear all content and formatting
  settingsSheet.setTabColor("#4a86e8"); // Blue tab color

  const settingsHeaders = ['Setting', 'Value'];
  const settingsData = [
    ['Gemini API Key', ''],
    ['TheirStack API Key', ''],
    ['Search Keywords', ''],
    ['Cover Letters Folder Link', ''],
    ['Resume Google Doc Link', ''],
    ['Writing Sample Links (Optional)', ''],
    ['Notification Email', '']
  ];

  // Delete extra rows and columns for a clean slate
  const requiredSettingsRows = settingsData.length + 1;
  if (settingsSheet.getMaxRows() > requiredSettingsRows) {
    settingsSheet.deleteRows(requiredSettingsRows + 1, settingsSheet.getMaxRows() - requiredSettingsRows);
  }
  if (settingsSheet.getMaxColumns() > 2) {
    settingsSheet.deleteColumns(3, settingsSheet.getMaxColumns() - 2);
  }

  // Add back rows if needed
  if (settingsSheet.getMaxRows() < requiredSettingsRows) {
    settingsSheet.insertRowsAfter(settingsSheet.getMaxRows(), requiredSettingsRows - settingsSheet.getMaxRows());
  }


  settingsSheet.getRange(1, 1, 1, 2).setValues([settingsHeaders]).setFontWeight('bold');
  settingsSheet.getRange(2, 1, settingsData.length, 1).setValues(settingsData.map(row => [row[0]]));
  
  settingsSheet.setColumnWidth(1, 250);
  settingsSheet.setColumnWidth(2, 400);

  const settingsRange = settingsSheet.getRange(1, 1, settingsData.length + 1, 2);
  settingsRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  

  // --- 4. Set up the Jobs Sheet ---
  jobsSheet.clear();
  jobsSheet.setTabColor("#6aa84f"); // Green tab color

  const jobsHeaders = ['Applied', 'Not Applying', 'Title', 'Pay', 'Link', 'Cover Letter', 'Hiring Manager', 'Posted Date', 'Employment Status', 'Industry', 'Sortable Pay', 'Job ID'];

  // Delete all columns beyond what we need
  if (jobsSheet.getMaxColumns() > jobsHeaders.length) {
    jobsSheet.deleteColumns(jobsHeaders.length + 1, jobsSheet.getMaxColumns() - jobsHeaders.length);
  }
  
  jobsSheet.getRange(1, 1, 1, jobsHeaders.length).setValues([jobsHeaders]).setFontWeight('bold');
  jobsSheet.setFrozenRows(1);

  jobsSheet.setColumnWidth(1, 60);
  jobsSheet.setColumnWidth(2, 90);
  jobsSheet.setColumnWidth(3, 350);
  jobsSheet.setColumnWidth(4, 200);
  jobsSheet.setColumnWidth(5, 100);
  jobsSheet.setColumnWidth(6, 120);
  jobsSheet.setColumnWidth(7, 150);
  jobsSheet.setColumnWidth(8, 100);
  jobsSheet.setColumnWidth(9, 120);
  jobsSheet.setColumnWidth(10, 120);
  jobsSheet.getRange('A1:L1').setBackground('#f3f3f3');


  // --- 5. Apply Formatting and Hide Columns ---
  setupConditionalFormatting(jobsSheet);
  jobsSheet.hideColumn(jobsSheet.getRange('K1')); // Hide Sortable Pay
  jobsSheet.hideColumn(jobsSheet.getRange('L1')); // Hide Job ID
  
  // --- 6. Lock the setup function ---
  properties.setProperty('SETUP_COMPLETE', 'true');


  // --- 7. Final Alert ---
  SpreadsheetApp.getUi().alert('Setup Complete!', 'Your "Jobs" and "Settings" sheets have been created and formatted. Please populate your details in the "Settings" tab to continue.', SpreadsheetApp.getUi().ButtonSet.OK);

  ss.setActiveSheet(jobsSheet);
}


/**
 * Sets up the conditional formatting rules for the Jobs sheet.
 * Green/Bold for 'Applied', Red/Strikethrough/Black Text for 'Not Applying'.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to apply formatting to.
 */
function setupConditionalFormatting(sheet) {
  sheet.setConditionalFormatRules([]); 
  // Apply to all columns including the hidden ones to be safe
  const range = sheet.getRange('A2:L' + sheet.getMaxRows());

  // Rule for 'Not Applying' (Checkbox in B is TRUE)
  const notApplyingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=($B2=TRUE)')
    .setBackground("#f4cccc") // Light red
    .setStrikethrough(true)
    .setFontColor("#000000") // Black text
    .setRanges([range])
    .build();

  // Rule for 'Applied' (Checkbox in A is TRUE)
  const appliedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=($A2=TRUE)')
    .setBackground("#d9ead3") // Light green
    .setBold(true)
    .setRanges([range])
    .build();
    
  const allRules = sheet.getConditionalFormatRules();
  allRules.push(appliedRule, notApplyingRule);
  sheet.setConditionalFormatRules(allRules);
}


/**
 * Fetches script settings from the 'Settings' sheet in a more robust way.
 * @returns {object} An object containing the settings.
 */
function getSettings() {
  const settingsSheet = SPREADSHEET.getSheetByName(SETTINGS_SHEET_NAME);
  const allSettings = settingsSheet.getDataRange().getValues();
  if (allSettings[0] && allSettings[0][0].toLowerCase().includes('setting')) allSettings.shift();
  
  const settingsMap = allSettings.reduce((acc, row) => {
    const key = row[0] ? row[0].trim() : '';
    const value = row[1] ? row[1].trim() : '';
    if (key) acc[key] = value;
    return acc;
  }, {});

  // More robustly find the API keys, ignoring small typos or variations.
  let theirStackApiKey = '';
  for (const key in settingsMap) {
      if (key.toLowerCase().includes('theirstack')) {
          theirStackApiKey = settingsMap[key];
      }
  }

  return {
    apiKey: settingsMap['Gemini API Key'],
    theirStackApiKey: theirStackApiKey,
    keywords: settingsMap['Search Keywords'],
    folderLink: settingsMap['Cover Letters Folder Link'],
    resumeDocLink: settingsMap['Resume Google Doc Link'],
    writingSampleLinks: settingsMap['Writing Sample Links'],
    notificationEmail: settingsMap['Notification Email']
  };
}

/**
 * Fetches all existing job IDs from the sheet to prevent duplicates.
 * @returns {Array<number>} An array of job IDs.
 */
function getExistingJobIds() {
  const sheet = SPREADSHEET.getSheetByName(JOBS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // Read from column L (the 12th column)
  const ids = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
  // Flatten array, convert to numbers, and remove empty/invalid values
  return ids.flat().map(id => parseInt(id, 10)).filter(id => !isNaN(id));
}

/**
 * Gets a Drive Folder object from a URL.
 * @param {string} url The URL of the Google Drive folder.
 * @returns {GoogleAppsScript.Drive.Folder} The Folder object.
 */
function getFolderFromUrl(url) {
  try {
    const folderIdMatch = url.match(/folders\/([a-zA-Z0-9-_]+)/);
    if (!folderIdMatch || !folderIdMatch[1]) {
      throw new Error("Could not extract a valid Folder ID from the URL.");
    }
    return DriveApp.getFolderById(folderIdMatch[1]);
  } catch (e) {
    Logger.log(`CRITICAL ERROR in getFolderFromUrl: ${e.message}`);
    throw new Error(`The provided folder link is invalid or you do not have permission to access it.`);
  }
}


/**
 * A helper function to get the text content from any Google Doc URL.
 * @param {string} url The URL of the Google Doc.
 * @returns {string} The text content of the document.
 */
function getTextFromDocUrl(url) {
  try {
    const docIdMatch = url.match(/d\/([a-zA-Z0-9-_]+)/);
    if (!docIdMatch) throw new Error(`Could not extract a valid Document ID from the URL: ${url}`);
    
    const exportUrl = `https://docs.google.com/document/d/${docIdMatch[1]}/export?format=txt`;
    const options = { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(exportUrl, options);
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to access Google Doc at ${url}. Code: ${response.getResponseCode()}.`);
    }
    return response.getContentText();

  } catch (e) {
    Logger.log(`CRITICAL ERROR in getTextFromDocUrl: ${e.message}`);
    throw e; // Re-throw the error to be caught by the main function
  }
}

/**
 * Gets the text content from the resume and extracts key details.
 * @param {string} url The URL of the resume Google Doc.
 * @returns {{fullName: string, phone: string, email: string, fullText: string}} An object with resume details.
 */
function getResumeDetails(url) {
  const fullText = getTextFromDocUrl(url);
  // Extract details using regex
  const name = fullText.split('\n')[0].trim(); // Assume name is the first line
  const phoneMatch = fullText.match(/\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/);
  const emailMatch = fullText.match(/[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}/);

  return {
    fullName: name || "Your Name",
    phone: phoneMatch ? phoneMatch[0] : "Your Phone",
    email: emailMatch ? emailMatch[0] : "Your Email",
    fullText: fullText
  };
}

/**
 * Fetches and concatenates the text from multiple writing sample documents.
 * @param {string} urlString A comma-separated string of Google Doc URLs.
 * @returns {string} The combined text of all documents.
 */
function getWritingSamplesText(urlString) {
  if (!urlString) return "";
  const urls = urlString.split(',').map(url => url.trim()).filter(Boolean);
  let combinedText = "";
  urls.forEach(url => {
    try {
      combinedText += getTextFromDocUrl(url) + "\n\n---\n\n";
    } catch (e) {
      Logger.log(`Skipping writing sample due to error: ${e.message}`);
    }
  });
  return combinedText;
}


/**
 * Main function to find and process new jobs.
 */
function findAndProcessNewJobs() {
  try {
    const settings = getSettings();
    if (!settings.apiKey || !settings.keywords) throw new Error("A Gemini API Key and Search Keywords are required.");
    if (!settings.theirStackApiKey) {
      throw new Error("You must provide an API key for TheirStack in the Settings sheet.");
    }
    
    SPREADSHEET.toast('Searching for jobs with TheirStack...');
    const newJobs = findJobsWithTheirStack(settings);

    if (newJobs.length > 0) {
      processJobs(newJobs);
      sortJobsSheet();
      sendNotificationEmail(newJobs.length, settings.notificationEmail);
      SPREADSHEET.toast(`Success! ${newJobs.length} new jobs were added.`);
    } else {
       Logger.log('No new jobs found during this run.');
       SPREADSHEET.toast('No new jobs were found that match your criteria.');
    }
  } catch (e) {
    Logger.log(`Error in findAndProcessNewJobs: ${e.message}`);
    SPREADSHEET.toast(`An error occurred: ${e.message}`);
  }
}

/**
 * Finds jobs using the TheirStack service.
 * @param {object} settings The script settings.
 * @returns {Array<object>} An array of job objects.
 */
function findJobsWithTheirStack(settings) {
    const url = `https://api.theirstack.com/v1/jobs/search`;
    const existingJobIds = getExistingJobIds();

    const payload = {
        include_total_results: false,
        posted_at_max_age_days: 30,
        job_country_code_or: ["US"],
        job_title_or: [settings.keywords],
        company_country_code_or: ["US"],
        min_salary_usd: 1,
        remote: true,
        page: 0,
        limit: 50,
        blur_company_data: false,
        job_id_not: existingJobIds
    };

    const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
            'Authorization': 'Bearer ' + settings.theirStackApiKey
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const responseText = response.getContentText();
        const result = JSON.parse(responseText);

        if (result.error) {
            const errorMessage = typeof result.error === 'object' ? JSON.stringify(result.error) : result.error;
            throw new Error(`TheirStack API returned an error: ${errorMessage}`);
        }
        
        if (!result.data || result.data.length === 0) {
            Logger.log('TheirStack found no new jobs.');
            return [];
        }

        // Map the response to our standard job object format
        const jobs = result.data.map(job => {
            const hiringManager = job.hiring_team?.[0];
            return {
                id: job.id,
                title: job.job_title,
                company: job.company_object?.name || "N/A",
                pay: job.salary_string || "Not specified",
                link: job.final_url || job.source_url || job.url,
                hiringManagerName: hiringManager?.full_name || null,
                hiringManagerLink: hiringManager?.linkedin_url || null,
                employment_status: job.employment_statuses?.[0] || 'N/A',
                industry: job.company_object?.industry || 'N/A',
                date_posted: job.date_posted ? new Date(job.date_posted).toLocaleDateString() : 'N/A'
            };
        });

        return jobs.filter(job => job.link); // Basic filter to ensure jobs have a link

    } catch (e) {
        Logger.log(`Error with TheirStack API: ${e.toString()}`);
        throw new Error("Failed to get jobs from TheirStack. Check logs for details.");
    }
}

/**
 * Parses a raw pay string into a numerical average hourly rate for sorting.
 * @param {string} payString The raw pay string.
 * @returns {number} The calculated average hourly rate.
 */
function parsePay(payString) {
  if (!payString || typeof payString !== 'string' || payString.toLowerCase() === 'not specified') return 0;
  
  let cleanPay = payString.toLowerCase().replace(/[$,]/g, '');
  cleanPay = cleanPay.replace(/k/g, '000');

  let numbers = (cleanPay.match(/\d+(\.\d+)?/g) || []).map(parseFloat);
  if (numbers.length === 0) return 0;
  
  let payValue = (numbers.length > 1) ? (numbers[0] + numbers[1]) / 2 : numbers[0];
  
  if (cleanPay.includes('year') || cleanPay.includes('annual')) {
      return payValue / 2080;
  }
  if (cleanPay.includes('month')) {
      return (payValue * 12) / 2080;
  }
  if (cleanPay.includes('hour')) {
      return payValue;
  }
  
  if (payValue > 1000) {
      return payValue / 2080;
  }
  
  return payValue;
}


/**
 * Sorts the 'Jobs' sheet based on status and pay.
 */
function sortJobsSheet() {
  const sheet = SPREADSHEET.getSheetByName(JOBS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort([
    { column: 2, ascending: true },  // "Not Applying" (FALSE first)
    { column: 1, ascending: true },  // "Applied" (FALSE first)
    { column: 11, ascending: false } // Sortable Pay (Column K, descending)
  ]);
  SpreadsheetApp.flush();
}

/**
 * Processes new jobs by adding them to the sheet with placeholders.
 */
function processJobs(newJobs) {
  const sheet = SPREADSHEET.getSheetByName(JOBS_SHEET_NAME);
  newJobs.forEach(job => {
    const displayTitle = job.company ? `${job.title} (${job.company})` : job.title;
    const formattedPay = job.pay && parsePay(job.pay) > 0 ? formatPay(job.pay) : "Not specified";
    const numericPay = parsePay(job.pay);
    const newRowIndex = sheet.getLastRow() + 1;

    const rowData = [false, false, displayTitle, formattedPay, 'Processing...', "Pending...", "", job.date_posted, job.employment_status, job.industry, numericPay, job.id];
    sheet.appendRow(rowData);
    
    const linkCell = sheet.getRange(newRowIndex, 5);
    if (job.link && job.link.startsWith('http')) {
      linkCell.setFormula(`=HYPERLINK("${job.link}", "link")`);
    } else {
      linkCell.setValue(job.link || 'N/A');
    }

    const hmCell = sheet.getRange(newRowIndex, 7); // Column G
    if (job.hiringManagerName && job.hiringManagerLink) {
        hmCell.setFormula(`=HYPERLINK("${job.hiringManagerLink}", "${job.hiringManagerName}")`);
    } else if (job.hiringManagerName) {
        hmCell.setValue(job.hiringManagerName);
    }

    sheet.getRange(newRowIndex, 6).setValue("Pending...");

    sheet.getRange(newRowIndex, 1, 1, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  });
}

/**
 * Finds the next job with a "Pending..." cover letter status, generates the letter,
 * and updates the sheet. Designed to be run on a time-based trigger.
 */
function generateSingleCoverLetter() {
  const sheet = SPREADSHEET.getSheetByName(JOBS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const coverLetterStatuses = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
  
  let pendingRowIndex = -1;
  for (let i = 0; i < coverLetterStatuses.length; i++) {
    if (coverLetterStatuses[i][0] === 'Pending...') {
      pendingRowIndex = i + 2;
      break;
    }
  }

  if (pendingRowIndex === -1) {
    Logger.log("No pending cover letters to generate.");
    return;
  }
  
  const coverLetterCell = sheet.getRange(pendingRowIndex, 6);
  coverLetterCell.setValue("Generating...");

  try {
    const settings = getSettings();
    const resumeDetails = getResumeDetails(settings.resumeDocLink);
    const writingSamples = getWritingSamplesText(settings.writingSampleLinks);
    
    const rowValues = sheet.getRange(pendingRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const linkFormula = sheet.getRange(pendingRowIndex, 5).getFormula();
    const linkMatch = linkFormula ? linkFormula.match(/=HYPERLINK\("([^"]+)"/) : null;
    
    const hmCell = sheet.getRange(pendingRowIndex, 7);
    const hmFormula = hmCell.getFormula();
    let hiringManagerName = hmCell.getValue();
    if (hmFormula) {
        const nameMatch = hmFormula.match(/, "([^"]+)"\)/);
        if (nameMatch) hiringManagerName = nameMatch[1];
    }


    const job = {
      title: rowValues[2].match(/(.*) \(/)?.[1] || rowValues[2],
      company: rowValues[2].match(/\(([^)]+)\)/)?.[1] || "N/A",
      pay: rowValues[3],
      link: linkMatch ? linkMatch[1] : '',
      hiringManagerName: hiringManagerName
    };

    const coverLetterUrl = generateCoverLetter(job, settings, resumeDetails, writingSamples);

    if (coverLetterUrl.startsWith('http')) {
      coverLetterCell.setFormula(`=HYPERLINK("${coverLetterUrl}", "Cover Letter")`);
    } else {
      coverLetterCell.setValue(coverLetterUrl);
    }
    Logger.log(`Successfully generated cover letter for row ${pendingRowIndex}`);

  } catch (e) {
    Logger.log(`Failed to create cover letter for row ${pendingRowIndex}: ${e.message}`);
    coverLetterCell.setValue('Error creating document');
  }
}


/**
 * Standardizes the pay string into a consistent "$XX.XX/hour ($YYY,YYY/year)" format.
 * @param {string} payString The raw pay string from the job posting.
 * @returns {string} The formatted pay string.
 */
function formatPay(payString) {
  const hourlyRate = parsePay(payString);
  if (hourlyRate === 0) return "Not specified";
  const annualRate = hourlyRate * 2080;
  const formattedHourly = hourlyRate.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
  const formattedAnnual = annualRate.toLocaleString('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0, maximumFractionDigits: 0 });
  return `${formattedHourly}/hour (${formattedAnnual}/year)`;
}


/**
 * Sends an email notification when new jobs are found.
 */
function sendNotificationEmail(jobCount, recipientEmail) {
  if (!recipientEmail) return;
  try {
    const subject = `[Job Tracker] ${jobCount} New Job${jobCount > 1 ? 's' : ''} Found!`;
    const body = `Hello,\n\n${jobCount} new job(s) have been added to your job tracker.\n\nView them here: ${SPREADSHEET.getUrl()}`;
    MailApp.sendEmail(recipientEmail, subject, body);
  } catch (e) {
    Logger.log(`Failed to send notification email: ${e.message}`);
  }
}

/**
 * Generates a cover letter using Gemini and saves it as a Google Doc.
 */
function generateCoverLetter(job, settings, resumeDetails, writingSamples) {
  const currentDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const jobTitleOnly = job.title.replace(/\s\(.*\)/, '');
  const salutation = job.hiringManagerName || "Hiring Manager";
  
  const prompt = `
You are an expert cover letter writer and a meticulous proofreader. Your primary goal is to generate a complete, professional, three-paragraph cover letter that is tailored to a specific job.

**Primary Instructions:**
1.  **Analyze Context:** First, carefully analyze the provided resume, the job details, and any writing samples to understand the candidate's skills, experience, and unique writing voice.
2.  **Identify Keywords:** Cross-reference the resume and the job details to identify key skills, qualifications, and keywords that overlap. These are the most important points to highlight.
3.  **Craft Unique Content:** Write three compelling and distinct body paragraphs. Do not use a generic template. The content should be unique to this specific application, persuasively arguing why the candidate's experience (from the resume) is a perfect match for the job's requirements (from the job details).
4.  **Emulate Writing Style:** If writing samples are provided, emulate the author's tone, style, and vocabulary. If not, use a standard professional and confident tone.
5.  **Follow Format and Proofread:** Adhere strictly to the provided template for the overall structure. Ensure the final output is grammatically perfect, free of spelling errors, and polished for professional use.

**Cover Letter Template:**

${resumeDetails.fullName}
${resumeDetails.phone}
${resumeDetails.email}

${currentDate}

${job.company}
Hiring Manager

Dear ${salutation},

(Paragraph 1: Introduction. State the position you are applying for and express genuine enthusiasm. Briefly introduce your key qualifications, connecting them directly to the job's core requirements.)

(Paragraph 2: Body. Provide a specific example from your resume that demonstrates your most relevant skills. Explain how this experience has prepared you for the responsibilities of this new role. Weave in the keywords you identified.)

(Paragraph 3: Closing. Reiterate your interest in the company and the role. Express your confidence in your ability to contribute to the team and include a call to action, such as your eagerness to discuss your qualifications further.)

Sincerely,
${resumeDetails.fullName}

---
**RESUME FOR CONTEXT:**
${resumeDetails.fullText}

---
**JOB DETAILS FOR CONTEXT:**
Title: ${job.title} (${job.company})
Pay: ${job.pay}

---
**WRITING STYLE SAMPLES FOR TONE AND STYLE ANALYSIS:**
${writingSamples}
`;

  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload) };
  const response = UrlFetchApp.fetch(GEMINI_API_ENDPOINT_GENERATE + settings.apiKey, options);
  const result = JSON.parse(response.getContentText());
  const coverLetterText = result.candidates[0].content.parts[0].text;
  const folder = getFolderFromUrl(settings.folderLink);
  const docName = `Cover Letter - ${job.title.replace(/[^a-zA-Z0-9\s]/g, "")}`;
  const doc = DocumentApp.create(docName);
  doc.getBody().setText(coverLetterText);
  doc.saveAndClose();
  const docFile = DriveApp.getFileById(doc.getId());
  folder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);
  return docFile.getUrl();
}


/**
 * DEBUGGING FUNCTION: A function to reset the setup lock, allowing runInitialSetup to be run again.
 */
function resetSetupLock() {
  try {
    PropertiesService.getScriptProperties().deleteProperty('SETUP_COMPLETE');
    SpreadsheetApp.getUi().alert('Success', 'The setup lock has been reset. You can now run the initial setup again. Warning: This will delete your existing data.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(`Error resetting setup lock: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', `Could not reset the setup lock. Error: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

<div align="center">
  <img src="https://ssl.gstatic.com/docs/script/images/logo/script-128.png" alt="Google Apps Script Logo" width="150">
</div>

<h1 align="center">AI Job Tracker & Cover Letter Generator</h1>

<p align="center">
  <img src="https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white" alt="Google Sheets Badge"/>
  <img src="https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white" alt="Apps Script Badge"/>
  <img src="https://img.shields.io/badge/Gemini%20API-8E77F0?style=for-the-badge&logo=google-gemini&logoColor=white" alt="Gemini API Badge"/>
  <img src="https://img.shields.io/badge/TheirStack%20API-FF4F8A?style=for-the-badge&logoColor=white&labelColor=1a1a1a&label=TheirStack" alt="TheirStack API Badge"/>
</p>

This repository contains the Google Apps Script code for an automated job tracking and cover letter generation system. It leverages the TheirStack API for job discovery, Google Sheets for tracking, and the Gemini API for AI-powered content creation, all running within a Google Workspace environment.

**Project Status:** `Active`
**Last Updated:** October 14, 2025

---

## 🚀 Key Features

-   **Automated Job Discovery:** Fetches job postings from the TheirStack API based on keywords, location, and other filters.
-   **AI-Powered Cover Letters:** Generates unique, professional cover letters for each job using Google's Gemini API, tailored to your resume, writing style, and the job description.
-   **Intelligent De-duplication:** Uses unique job IDs to ensure you never fetch the same job twice, saving API credits and keeping your list clean.
-   **Background Processing:** Intelligently generates cover letters one at a time using time-based triggers to work around Google's script execution time limits.
-   **Dynamic Workflow:** Use simple checkboxes to mark jobs as "Applied" or "Not Applying." The sheet automatically re-sorts itself, and cover letter files are managed in Google Drive accordingly.
-   **Personalized Voice & Tone:** Analyzes your own writing samples (from Google Docs) to learn your unique style, ensuring the AI-generated cover letters sound authentically like you.
-   **Automated Sheet Setup:** A one-click setup function creates, formats, and protects all necessary sheets and columns.

---

## 🛠️ Technology Stack

This project is built entirely within the Google Cloud ecosystem, requiring no external servers or hosting.

- **Platform:** [Google Apps Script](https://developers.google.com/apps-script)
- **Database/UI:** [Google Sheets](https://www.google.com/sheets/about/)
- **Job Discovery:** [TheirStack API](https://theirstack.com/)
- **Cover Letter Generation:** [Google Gemini API](https://ai.google.dev/)
- **File Storage:** [Google Drive](https://www.google.com/drive/)

---

## ▶️ Getting Started / Installation

To set up your own instance of this job tracker, follow these steps.

1.  **Create a Blank Google Sheet:**
    Start by creating a new, empty spreadsheet at [sheet.new](https://sheet.new).

2.  **Open the Apps Script Editor:**
    In your new sheet, navigate to `Extensions` > `Apps Script`. This will open the editor in a new tab.

3.  **Paste the Script:**
    Delete any boilerplate code in the `Code.gs` file. Copy the entire contents of the `Code.gs` file from this repository and paste it into the editor. Click the **Save project** icon.

4.  **Run Initial Setup:**
    Return to your Google Sheet (you may need to refresh the page). A new **"Job Automator"** menu will appear. Click `Job Automator` > `1. Run Initial Setup (Formatting)`.

5.  **Authorize the Script:**
    A pop-up will ask for permission. Click `Continue`, select your Google account, click `Advanced`, then **"Go to... (unsafe)"**, and finally `Allow`. This is required for the script to manage your sheets and files.

6.  **Configure Your Settings:**
    The setup process creates a `Settings` sheet. Gather your API keys and links and paste them into the appropriate `Value` cells.

7.  **Set Automated Triggers:**
    This is the final step to enable automation. Go back to the Apps Script editor and click the **Triggers** (clock) icon. Click **+ Add Trigger** and create the three required triggers as detailed in the `Triggers` section below.

---

## ⚙️ Configuration & Triggers

The script is configured via the `Settings` sheet and relies on three triggers for full automation.

### Settings Sheet

| Setting                       | Description                                                                                                        |
| ----------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| **Gemini API Key** | Your API key from [Google AI Studio](https://aistudio.google.com/app/apikey) for generating cover letters.             |
| **TheirStack API Key** | Your API key from [TheirStack.com](https://theirstack.com/) for finding jobs.                                        |
| **Search Keywords** | The job title or keywords to search for (e.g., "Technical Writer", "Remote Software Engineer").                    |
| **Cover Letters Folder Link** | The URL of the Google Drive folder where cover letters will be saved.                                              |
| **Resume Google Doc Link** | A shareable link to your resume in Google Docs format (must be set to "Anyone with the link can view").             |
| **Writing Sample Links** | (Optional) Comma-separated links to Google Docs containing samples of your writing to train the AI on your tone. |
| **Notification Email** | The email address where you want to receive notifications when new jobs are found.                                 |

### Triggers

Three triggers must be set in the Apps Script editor for the system to run automatically:

1.  **`findAndProcessNewJobs`**
    - **Event:** `Time-driven`
    - **Frequency:** `Day timer` (e.g., every morning)
    - **Purpose:** Searches for new jobs.

2.  **`generateSingleCoverLetter`**
    - **Event:** `Time-driven`
    - **Frequency:** `Minutes timer` (e.g., every 1 or 5 minutes)
    - **Purpose:** Processes the queue of pending cover letters in the background.

3.  **`handleEdit`**
    - **Event:** `From spreadsheet` > `On edit`
    - **Purpose:** Automatically sorts the sheet and manages files when a checkbox is changed.

---

## 📂 Project Components

The project is self-contained within a single Google Sheet and its bound Apps Script file.

```text
Job Tracker (Google Sheet)
│
├── Jobs (Sheet)              # The main dashboard for tracking job applications.
│
└── Settings (Sheet)          # Configuration panel for API keys, links, and keywords.
│
└── App Script: Code.gs       # The single script file containing all automation logic.
    │
    ├── onOpen()              # Creates the custom "Job Automator" menu in the UI.
    ├── runInitialSetup()     # Creates and formats the sheets, sets up locks.
    ├── findAndProcessNewJobs()# Main function to call the TheirStack API and add jobs.
    ├── generateSingleCoverLetter()# AI function to generate one cover letter per run.
    └── handleEdit()          # Triggered function that responds to user edits.
```

---

## 💡 Troubleshooting & FAQ

-   **Q: Why am I getting a "timeout" error?**
    -   **A:** Google Apps Script has a 6-minute execution limit. This project is specifically designed to avoid this by generating cover letters one by one in the background. If you see this error, it's likely during the initial setup of a very large list. Just run the "Generate Next Cover Letter" function a few times manually to process the queue.

-   **Q: No jobs are being found. Why?**
    -   **A:** Double-check that your TheirStack API key and Search Keywords are correctly entered in the `Settings` sheet. Also, ensure your TheirStack account is active and has available API credits.

-   **Q: How do I view error logs?**
    -   **A:** In the Apps Script editor, click the **Executions** (list) icon on the left. Find the failed run and click on it to see the detailed logs, which will contain any error messages.

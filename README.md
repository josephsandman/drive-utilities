# Google Apps Script Utilities

## Overview

This repository contains Google Apps Script utilities for managing Google Sheets and Gmail interactions, with capabilities for performing tasks like mail merges, copying files and folders, and retrieving metadata from Google Drive. The script interacts with Google Sheets and Google Forms to automate email communications effectively.

## Table of Contents
- [Features](#features)
- [Getting Started](#getting-started)
- [Usage](#usage)
- [Function Documentation](#function-documentation)
- [Contributing](#contributing)
- [License](#license)

## Features
- Custom menus for easy access to utilities within Google Sheets.
- Mail merge functionality utilizing Gmail drafts.
- Automated copying of files and folders based on user input.
- Query removal from Google Drive URLs.
- Error handling and logging throughout the script execution.

## Getting Started

### Prerequisites
- Access to Google Workspace with permission to use Google Sheets, Forms, and Drive.
- A Google account with a project created in Google Apps Script.

### Setup Instructions
1. **Clone the Repository**:
   ```bash
   git clone https://github.com/yourusername/repo-name.git
   cd repo-name
   ```

2. **Open Google Apps Script**:
   - Go to [Google Apps Script](https://script.google.com/) and create a new project.

3. **Copy Code**:
   - Copy the contents from your cloned repository files into the new Apps Script project.

4. **Deploy**:
   - Set triggers like `onOpen` to initialize the custom menus.
   - Authorize the script to access your Google account scopes when prompted.

## Usage

1. **Open your Google Sheet**:
   - Upon opening, the custom menus will be available in the menu bar.
   
2. **Access Utilities**:
   - Navigate to the custom menus created (e.g., *Drive utilities* and *Gmail utilities*).
   - Select the desired option to prompt user interactions for executing script functionalities.

3. **Mail Merge Example**:
   - When executing the mail merge functionality, you will be prompted for:
     - Subject line for emails.
     - Sheets and header names used for recipient addresses and sent status.

## Function Documentation

- **`sendEmails(subjectLine, thisSheet, thisTab, emailRecipients, emailSent)`**: Sends emails based on recipient data in the provided Google Sheet. If headers are missing, prompts for user input.

- **`retrieveFiles()`**: Retrieves file names and URLs from a specified Google Drive folder and writes them to the active sheet.

- **`retrieveFolders()`**: Similar to `retrieveFiles`, but for subfolders within a specified Google Drive folder.

- **`getUserInput(promptMessage)`**: Displays a prompt to get user input.

- **`removeQueryFromUrl(url)`**: Removes query strings from a Google Drive URL.

- **`getIdFromUrl(url)`**: Extracts the folder ID from a Google Drive URL.

## Contributing

We welcome contributions! Please read our [Contributing Guidelines](CONTRIBUTING.md) for details on our code of conduct and the process for submitting pull requests to our repository.

## License

This project is licensed under the Apache License 2.0. See [LICENSE](LICENSE) for more details.

---

### Note
Make sure to replace `yourusername` and `repo-name` with your actual GitHub username and the name of your repository.

Feel free to enhance the README with additional sections, such as FAQs or a changelog, based on the specific needs of your project or potential users. This document should help new developers quickly understand how to set up and use the codebase effectively.
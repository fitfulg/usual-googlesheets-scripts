# Google Sheets Scripts

Personal repo for my own customized "Google - Apps Scripts" code to enhance Google Sheets with additional formatting, validation, and chart creation capabilities.

The initial idea was just to practice the integration from this repo up to Google Apps Script via GitHub Actions, but later I decided to also publish the small improvements that I am making in my custom template for the TODO sheet that I use on a daily basis.

Apps Script uses Javascript and runs on Google Cloud.

## Setup Instructions

1. **Clone the repository to your local machine:**
    ```bash
    git clone https://github.com/yourusername/myGoogleSheetsScripts.git
    cd myGoogleSheetsScripts
    ```

2. **Install Node.js and clasp:**
    Make sure you have Node.js installed. If not, you can download and install it from [Node.js official website](https://nodejs.org/).

    Then, install `clasp` globally:
    ```bash
    npm install -g @google/clasp
    ```

3. **Authenticate with Google:**
    ```bash
    clasp login
    ```

4. **Create a new Google Apps Script project or use an existing one:**
    - To create a new project:
        ```bash
        clasp create --type sheets --title "My Project"
        ```
    - To use an existing project, get the script ID from the Apps Script URL and set it in `.clasp.json`:
        ```json
        {
          "scriptId": "YOUR_SCRIPT_ID",
          "rootDir": "./"
        }
        ```

5. **Push the code to Google Apps Script:**
    ```bash
    clasp push
    ```

## Using GitHub Actions for CI/CD

This project is set up to use GitHub Actions for continuous deployment to Google Apps Script. The `setup-clasp.yml` workflow takes care of the deployment process.

### Setup GitHub Secrets

You need to set up the following secrets in your GitHub repository:

- `CLASP_TOKEN`: This is your `clasp` authentication token. You can get it by running `clasp login --no-localhost` and copying the token from the URL.
- `CLASP_SCRIPT_ID`: This is the script ID of your Google Apps Script project.

## Usage

After setting up the scripts, a custom menu named **Custom Formats** will appear in your Google Sheets document. You can use this menu to apply custom formats, background colors, and create pie charts based on the data in your sheet.
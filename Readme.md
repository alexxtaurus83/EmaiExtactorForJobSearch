# Job Application Email Parser

This is a .NET console application designed to automatically parse emails related to job applications from a specified Gmail label. It extracts key information such as company name, job title, and application status, then compiles this data into a structured CSV file for easy tracking and analysis.

***

> **Note from the Developer:** This project was developed in collaboration with Google's Gemini. The AI assisted in writing, refactoring, and debugging the C# code, as well as externalizing the configuration and generating this documentation. This collaboration significantly accelerated the development process and improved code quality.

***

## Features

- **IMAP Email Fetching**: Securely connects to a Gmail account using IMAP to read emails.
- **Intelligent Data Extraction**: Uses a sophisticated set of configurable regular expressions and keyword matching to parse the following from email subjects and bodies:
    - Company Name
    - Job Title
    - Application Status (`Original Application`, `Rejection Reply`, `Other`)
    - Direct Company Flag (identifies if the email is directly from the company or a third party).
- **External Configuration**: All settings, credentials, keywords, and regex patterns are stored in an `appsettings.json` file, allowing for easy updates without recompiling the code.
- **CSV Export**: Outputs all extracted data into a clean, easy-to-read CSV file, ordered by date.

## How to Use

### Setup and Execution

The application will connect to your email, fetch the messages, process them, and create a `simplified_emails.csv` file in the same directory.

## Configuration (`appsettings.json`)

The `appsettings.json` file gives you full control over the parsing logic.

- **`imapSettings`**: Contains your credentials and server details.
- **`extractionPatterns`**: Holds the regex patterns used to find company names and job titles. You can add or modify these to improve accuracy.
- **`domainLists`**: Lists of domains to help classify emails (e.g., distinguishing an Applicant Tracking System (ATS) from a generic email provider).
- **`keywordLists`**:
    - `specificCompanyKeywords`: Helps identify companies that are hard to catch with regex.
    - `rejectionKeywords` & `originalApplicationKeywords`: Used to classify the application status.
- **`outputSettings`**: Defines the name of the output CSV file.
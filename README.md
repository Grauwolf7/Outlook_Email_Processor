# Outlook Email Summarizer

Transform your email management with this intelligent Python tool that bridges Microsoft Outlook and AI summarization capabilities. This application helps you stay on top of your inbox by automatically organizing and preparing your emails for efficient processing.

## What Makes This Tool Special

Have you ever wished you could quickly understand what happened in your inbox over the past few days? This tool does exactly that by collecting your Outlook emails and organizing them in a way that's perfect for AI summarization. Think of it as your personal email librarian that carefully catalogs and prepares your messages for review.

## How It Works

The tool works as follows:

1. **Email Collection**: Connects to your active Outlook instance and retrieves emails from your inbox based on the date range you specify.

2. **Smart Organization**: Creates a structured folder system organized by date, making it easy to track and manage your email history.

3. **Data Processing**: Converts your emails into a clean JSON format while preserving all important details like sender information, timestamps, and message content.

4. **AI Preparation**: Generates carefully formatted prompts that you can use with ChatGPT or similar AI tools to create summaries of your email communications.

## Getting Started

Before you begin, make sure you have:
- Windows operating system
- Microsoft Outlook installed and running
- Python 3.x installed on your system

To install the tool:

### Clone the repository
```git clone https://github.com/YourUsername/OutlookEmailSummary.git```

### Install required Python package
```pip install pywin32```

## Using the Tool
Running the tool is straightforward:

1. Open your terminal and navigate to the project directory
2. Run the script:
```python email_summarizer.py```
3. When prompted, enter how many days back you want to process:
   -  Enter 0 to process today's emails only
   -  Enter any positive number to process that many days into the past

The tool will create an organized structure in an "Email Summaries" folder:

```
Email Summaries/
└── 2024-01-28/
    ├── emails.json           # Structured email data
    └── email_summary_prompt.txt  # Ready-to-use AI prompt
```

## Customization Options

You can tailor the tool to your needs by:

1. Email Filtering: Add email addresses to the excluded_addresses set in the should_include_email function to skip certain senders.

3. Date Processing: Modify the date range processing by adjusting the loop in the main function.

4. Prompt Format: Customize the AI prompt generation in the create_summary_prompt function to better suit your summarization needs.

## Understanding the Output
The tool generates two key files for each day:

1. emails.json: A structured JSON file containing:
  
    - Email subjects
    - Sender information
    - Timestamps
    - Message content


2. email_summary_prompt.txt: A formatted prompt that:

    - Groups emails by date
    - Includes all relevant email details
    - Is optimized for AI summarization

## Technical Details
This tool leverages several key technologies:

  - win32com.client for Outlook integration
  - Python's pathlib for robust file system operations
  - JSON handling for structured data storage
  - DateTime manipulation for precise email filtering

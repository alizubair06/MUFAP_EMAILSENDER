# MUFAP AL Habib Mutual Funds Scraper & Automation

Automated daily data scraping pipeline for **AL Habib Mutual Funds** performance metrics from the official Mutual Funds Association of Pakistan ([MUFAP](https://mufap.com.pk/Industry/IndustryStatDaily?tab=1)) portal.

## Features
- **Categorized HTML Email**: Preserves Sector/Category hierarchy headings in the email body.
- **Excel Export (`.xlsx`)**: Attached workbook containing structured data for search, sorting, and offline analysis.
- **Automated Polling Window**: Scheduled via GitHub Actions to poll every 20 minutes from **9:00 AM PKT to 8:00 PM PKT** on weekdays.
- **Duplicate Prevention**: Updates `state.json` upon sending today's report, preventing repeat emails when MUFAP updates.

## Setup Instructions

### 1. GitHub Secrets
Add the following secrets to your GitHub Repository (**Settings > Secrets and variables > Actions**):

| Secret Name | Example Value |
| :--- | :--- |
| `SMTP_SERVER` | `smtp.gmail.com` |
| `SMTP_PORT` | `587` |
| `SENDER_EMAIL` | `your_email@gmail.com` |
| `SENDER_PASSWORD` | `xxxx xxxx xxxx xxxx` (App Password) |
| `RECIPIENT_EMAILS` | `person1@gmail.com,person2@gmail.com` |

### 2. Local Execution
```bash
# Clone Repository

# Install dependencies
pip install -r requirements.txt

# Run script
python main.py
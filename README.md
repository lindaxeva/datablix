# Datablix 1.0

Datablix is a data quality and verification assistant for research-based directory projects.

It transforms uploaded research spreadsheets into structured, review-ready directories by identifying missing information, possible duplicates, invalid source links, questionable dates, and records requiring manual verification.

## Version 1 MVP

Version 1 includes:

* Streamlit welcome page
* CSV and Excel file upload
* Data preview
* Standard Datablix directory columns
* Missing-field checks
* Automated QA flags
* Data quality KPI cards
* Editable manual review queue
* Review status and reviewer notes
* CSV result downloads
* Fictional sample data for testing

## Standard Datablix Fields

The standard directory structure includes:

* Record ID
* Name
* Category
* Address
* City
* Province
* Postal Code
* Phone
* Email
* Website
* Source URL
* Date Researched
* Verification Status
* Reviewer Notes

## Current QA Checks

Datablix currently checks for:

* Missing required fields
* Missing standard columns
* Duplicate Name and City combinations
* Invalid source URL formats
* Invalid research dates
* Future research dates
* Unrecognized verification statuses

## Running the App

Install the required packages:

```bash
pip install -r requirements.txt
```

Run the Streamlit application:

```bash
streamlit run app.py
```

## Privacy

Only fictional sample data should be used while building and testing this public MVP.

Confidential user information and private research data must not be uploaded to the public application or committed to this repository.

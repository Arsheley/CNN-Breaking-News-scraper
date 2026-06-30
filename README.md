# CNN Breaking News Scraper

A Python automation project that retrieves CNN breaking news through an API, prevents duplicate entries, and exports the results to a professionally formatted Excel workbook.

---

## Project Overview

This project automatically retrieves breaking news from CNN, extracts the required information, checks for duplicate headlines, and stores only new articles in an Excel workbook. The project demonstrates Python automation, API integration, data processing, and Excel automation.

---

## Features

- Retrieves breaking news from the CNN API.
- Stores headlines, publication date, article body, and scrape time.
- Prevents duplicate headlines from being saved.
- Automatically formats the Excel worksheet.
- Handles network and request errors gracefully.

---

## Technologies Used

- Python
- Requests
- OpenPyXL
- JSON
- Git
- GitHub

---

## How the Project Works

1. Sends a request to the CNN API.
2. Receives the response in JSON format.
3. Extracts the required article information.
4. Compares each headline with the existing headlines in the Excel workbook.
5. Saves only new articles.
6. Formats the worksheet automatically.
7. Saves the updated workbook.

---

## Challenges Faced

During development, several real-world challenges were encountered:

- Investigated the CNN website using the browser's Developer Tools.
- Inspected XHR network requests to identify the correct API endpoint.
- Diagnosed and resolved HTTP 400 Bad Request errors.
- Adapted the project after changes to the CNN API request.
- Implemented duplicate detection to avoid storing repeated news articles.

---

## What I Learned

This project strengthened my understanding of:

- Working with REST APIs.
- Parsing and processing JSON data.
- Excel automation using OpenPyXL.
- Error handling using `try` and `except`.
- Organizing Python programs using functions.
- Version control with Git.
- Publishing projects on GitHub.
- Debugging real-world software problems.

---

## Future Improvements

- Automatically detect updated API endpoints.
- Add logging for easier debugging.
- Export data to additional formats.
- Build a graphical user interface (GUI).
- Schedule the scraper to run automatically.

---

## Author

**Arsheley Mapesa**

Electrical and Electronic Engineering Student

**Interested in:**

- Python Automation


*Notes:*
- The **src** directory contains all source code.
- The **ui** folder within **src** is dedicated to user interface components.
- The **config** folder holds JSON configuration files, making them easily accessible and modifiable.
- The **tests** folder contains unit tests to validate functionality.
- The **requirements.txt** file will list dependencies for easy setup.

---

## Step-by-Step Development Plan

### Step 1: Project Initialization
- Set up a new Git repository.
- Create a virtual environment and install dependencies listed in `requirements.txt`.
- Set up the suggested file structure.

### Step 2: Develop Email Monitoring Module
- **Objective:** Monitor the Outlook client for new emails.
- Implement an event-driven or polling mechanism in `email_monitor.py` to detect incoming emails.
- Write initial tests in `test_email_parser.py` to ensure connectivity with Outlook.

### Step 3: Implement Email Parsing and Keyword Matching
- **Objective:** Extract email details and match content against keywords.
- Implement parsing functions in `email_parser.py` using data from `pre_defined_keywords.json`.
- Validate with sample emails and unit tests.

### Step 4: Create Excel Population Module
- **Objective:** Append email data into an Excel file based on configuration.
- Develop functions in `excel_handler.py` to load the Excel file and append data.
- Use `excel_sheet_mapping.json` to map fields to Excel columns.
- Test the functionality with dummy data.

### Step 5: Develop JSON Administration Functions & UI
- **Objective:** Create backend functions and a UI for managing JSON configuration files.
- In `json_admin.py`, create functions to load, validate, and update:
  - `company_name.json`
  - `incident_ref_code.json`
  - `pre_defined_keywords.json`
  - `excel_sheet_mapping.json`
- Develop corresponding UI screens in the **ui** directory for administration tasks.

### Step 6: Integrate Email Filtering by Date and Time
- **Objective:** Allow users to filter emails by a start date/time.
- Add UI elements in `main_window.py` for date/time input and "Apply Filter" button.
- Modify email processing functions to filter emails accordingly.

### Step 7: Implement Data Population Options in UI
- **Objective:** Provide both auto and manual data population options.
- Update the main window UI:
  - Include a checkbox for "Auto-Populate" mode.
  - Add a "Populate Email" button for manual processing.
- Ensure the UI displays processing status and updates the log.

### Step 8: Build Main UI & Workflow Integration
- **Objective:** Integrate all functionalities into a comprehensive UI.
- Combine all UI components in `main_window.py` and `settings.py`.
- Implement navigation between main and settings views.
- Incorporate feedback and status panels.

### Step 9: Error Handling and Logging
- Implement error handling in all modules using try/except blocks.
- Develop logging functionality in `logger.py` to capture errors and process summaries.
- Optionally, integrate a log viewer into the UI.

### Step 10: Testing and Validation
- Write unit tests for individual modules in the **tests** folder.
- Conduct integration testing with Outlook running.
- Validate that data mapping adheres to the configuration in `excel_sheet_mapping.json`.
- Ensure UI components work as expected.

### Step 11: Documentation and Final Adjustments
- Update internal code comments and documentation.
- Prepare user documentation for configuring and using the application.
- Perform final testing, debugging, and prepare the app for deployment.

---

## Functionality Details

### Outlook Email Monitoring and Processing
- **Monitoring:** The app must actively monitor the Outlook client.
- **Email Processing:**
  - Read sender, date received, and email content.
  - Check the email body against keywords defined in `pre_defined_keywords.json`.
  - If keywords are present, extract and log:
    - Email address (sender)
    - Date received
    - Company name (as mapped or extracted)
    - Incident reference code
- **Excel Population:**
  - Use `excel_sheet_mapping.json` to map extracted data to Excel columns.
  - Append data to the `Health check details` tab.

### JSON Administrative Functions & UI
- **Configuration Files:**
  - `company_name.json`
  - `incident_ref_code.json`
  - `pre_defined_keywords.json`
  - `excel_sheet_mapping.json`
- **UI Requirements:**
  - Display current configurations.
  - Allow adding, editing, and deleting entries.
  - Validate JSON structure on save.

### Email Filtering by Date and Time
- **Feature:**
  - Allow the user to specify a start date and time.
  - Filter emails to process only those received after the specified timestamp.
- **UI:** Provide date/time input and an "Apply Filter" button.

### Data Population Options
- **Auto-Populate:**
  - Checkbox in the main UI to enable real-time data population.
- **Manual Populate:**
  - A "Populate Email" button to trigger processing on demand.
- **User Feedback:**
  - Display confirmation messages and processing status in the UI.

---

## Error Handling and Logging

- **Error Handling:**
  - Implement error management in modules for email access, JSON I/O, and Excel file updates.
  - Display error messages in a user-friendly manner.
- **Logging:**
  - Log errors and process summaries using `logger.py`.
  - Optionally integrate a log viewer in the UI.

---

## Testing and Validation

- **Unit Tests:**
  - Develop tests for email parsing, keyword matching, JSON administration, and Excel population.
- **Integration Testing:**
  - Test complete workflows with the Outlook client active.
  - Ensure data mapping is correct per `excel_sheet_mapping.json`.
- **UI Testing:**
  - Validate that all UI components function as intended and that configurations persist.

---

## Additional Considerations

- **Concurrency:** Ensure the app handles multiple simultaneous email events without data loss.
- **User Notifications:** Implement notifications or visual cues for successful operations and errors.
- **Extensibility:** Design the architecture to allow easy addition of new functionalities in the future.

---

## Conclusion

Follow these detailed instructions to build the desktop Python application. Start with the basic email monitoring and parsing, then incrementally implement Excel population, JSON administration, and UI components. Testing at each step is essential to ensure robust functionality and a smooth user experience.

If any part of these instructions requires further clarification or modification, please confirm before proceeding.

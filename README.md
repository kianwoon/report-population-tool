# Report Population Tool

A tool for monitoring emails, extracting relevant information, and automatically populating Excel reports with the extracted data.

## Features

- **Email Monitoring**: Automatically monitors Outlook emails for new messages
- **Email Parsing**: Extracts structured data from emails using pattern matching and keyword recognition
- **Excel Integration**: Populates Excel reports with data extracted from emails
- **Configuration Management**: UI for managing company names, incident codes, and predefined keywords
- **User-friendly Interface**: Simple and intuitive UI for monitoring and administration

## Project Structure

```
report_population_tool/
├── src/                # Source code
│   ├── ui/             # User interface components
│   ├── utils/          # Utility functions
│   ├── email_monitor.py # Email monitoring functionality
│   ├── email_parser.py  # Email parsing functionality
│   ├── excel_handler.py # Excel handling functionality
│   ├── json_admin.py    # JSON configuration management
│   ├── admin_ui.py      # Administration UI
│   └── main.py          # Main application entry point
├── config/             # Configuration files
│   ├── company_name.json        # Company names configuration
│   ├── incident_ref_code.json   # Incident reference codes
│   ├── pre_defined_keywords.json # Predefined keywords for parsing
│   └── excel_sheet_mapping.json # Excel sheet mapping configuration
├── logs/               # Application logs
├── data/               # Data files (Excel reports)
├── tests/              # Unit tests
└── requirements.txt    # Project dependencies
```

## Setup

1. Create a virtual environment:
   ```
   python -m venv venv
   ```

2. Activate the virtual environment:
   - On macOS/Linux:
     ```
     source venv/bin/activate
     ```
   - On Windows:
     ```
     venv\Scripts\activate
     ```

3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

### Running the Main Application

```
python src/main.py
```

### Running the Admin UI

```
python src/main.py --admin
```

### Command Line Options

- `--admin`: Launch the administration UI for managing configurations
- `--debug`: Enable debug logging

## Configuration

The application uses JSON configuration files stored in the `config/` directory:

- `company_name.json`: List of company names for matching in emails
- `incident_ref_code.json`: Incident reference codes and their descriptions
- `pre_defined_keywords.json`: Keywords organized by categories for email parsing
- `excel_sheet_mapping.json`: Mapping of data fields to Excel columns

These configurations can be managed through the Admin UI.

## Development

This project follows clean code principles with a focus on:
- Simple, readable code
- Focused functionality
- Clear, consistent naming
- Core functionality before optimization

### Adding New Features

When adding new features:
1. Create focused modules with clear responsibilities
2. Add appropriate error handling and logging
3. Update the UI to expose the new functionality
4. Document the changes in comments and update this README

## Testing

Run tests with:
```
pytest tests/
```

## Troubleshooting

Common issues:

1. **Email monitoring not starting**: Ensure Outlook is properly configured and running
2. **Excel file not updating**: Check file permissions and ensure the file is not open in another application
3. **Configuration not loading**: Verify JSON files in the config directory are valid

For more detailed logs, check the `logs/` directory.

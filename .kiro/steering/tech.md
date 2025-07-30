# Technology Stack

## Core Technologies
- **Python 3.x**: Main programming language
- **tkinter + ttkbootstrap**: Modern GUI framework with themed interface
- **Selenium WebDriver**: Web automation for timesheet form filling
- **pandas + openpyxl**: Excel file manipulation and data processing
- **Google APIs**: Calendar integration (google-api-python-client, google-auth-*)

## Key Dependencies
```bash
pip install pandas selenium webdriver-manager openpyxl google-api-python-client google-auth-httplib2 google-auth-oauthlib ttkbootstrap
```

## Architecture Patterns
- **Modular Design**: Separate modules for GUI, automation, and calendar integration
- **Threading**: Background operations to prevent GUI freezing
- **Configuration-based**: Credentials and paths stored in `config.py`
- **Event-driven GUI**: Step-by-step workflow with state management

## File Structure Conventions
- `config.py`: User credentials and file paths (not committed)
- `config_example.py`: Template for configuration
- `timesheet_gui.py`: Main GUI application
- `timesheet_filler.py`: Selenium automation logic
- `google_calendar_integration.py`: Calendar API integration
- `working/`: Development and backup versions

## Common Commands
- **Run GUI**: `python timesheet_gui.py`
- **Run automation directly**: `python timesheet_filler.py [--headless] [--dry-run]`
- **Update from calendar**: `python update_calendar.py`

## Data Format
Excel files use Hebrew column headers:
- `שנה` (Year), `חודש` (Month), `יום` (Day)
- `זמן התחלה` (Start Time), `זמן סיום` (End Time), `שעות` (Hours)
- `מה` (Notes/Description)
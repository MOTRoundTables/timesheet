# Project Structure

## Root Directory Layout
```
timesheet/
├── config.py                    # User credentials and paths (gitignored)
├── config_example.py            # Configuration template
├── credentials.json             # Google API credentials (gitignored)
├── token.json                   # Google OAuth token (gitignored)
├── README.md                    # Comprehensive documentation
├── example.xlsx                 # Sample Excel file
├── ‏‏‏‏test2.xlsx                # Current working Excel file
├── timesheet_gui.py             # Main GUI application entry point
├── timesheet_filler.py          # Selenium automation core logic
├── google_calendar_integration.py # Google Calendar API integration
├── update_calendar.py           # Standalone calendar update script
├── working/                     # Development versions and backups
│   ├── google_ver1/            # Google integration development
│   └── *.py                    # Various working copies
├── build/                      # Build artifacts
├── dist/                       # Distribution files
└── __pycache__/               # Python bytecode cache
```

## Code Organization Principles
- **Single Responsibility**: Each module handles one main concern
- **Configuration Separation**: Sensitive data isolated in config files
- **Working Directory**: Keep development versions separate from production
- **Backup Strategy**: Automatic timestamped backups before modifications

## File Naming Conventions
- Main modules: descriptive names (e.g., `timesheet_gui.py`)
- Working copies: append descriptive suffixes (e.g., `timesheet_filler copy.py`)
- Configuration: `config.py` for active, `config_example.py` for template
- Data files: Hebrew filenames supported for Excel files

## Import Structure
- GUI imports automation and calendar modules
- Modules import `config` for shared settings
- Threading used for long-running operations
- Error handling with user-friendly messages

## Security Considerations
- Credentials stored in separate config file
- Google API tokens managed automatically
- Backup creation before data modifications
- Input validation for user-entered data
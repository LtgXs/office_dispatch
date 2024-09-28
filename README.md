# Office Dispatch

OfficeDispatch is an automated python script for detecting and uploading PowerPoint, Excel, and Word files to GitHub.

## Features

- **Configuration Loading and Validation**: Loads configuration from `config.json`. If the file does not exist or the configuration is invalid, it uses default settings.
- **File Processing**: Detects and copies new PowerPoint, Excel, and Word files to the target folder.
- **GitHub Upload**: Uploads new files to the specified GitHub repository.

## Installation

1. Clone this repository:
    ```bash
    git clone https://github.com/LtgXs/office_dispatch.git
    cd your_repo_name
    ```

2. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

3. Configure the `config.json` file:
    ```json
    {
        "repo_name": "your_repo_name",
        "github_token": "enter_your_github_token_here",
        "retry_interval": 10,
        "check_interval": 30
    }
    ```

## Usage

1. Run the script:
    ```bash
    python Checker.py
    ```

2. The script will automatically detect new PowerPoint, Excel, and Word files and upload them to GitHub.

## Configuration File Specification

- `repo_name`: The name of the GitHub repository.
- `github_token`: The GitHub access token.
- `retry_interval`: The retry interval (in seconds) when initializing COM objects fails.
- `check_interval`: The interval (in seconds) for detecting new files.

## Logging

All log messages are written to the `log.txt` file, including timestamps and detailed information.

## Contributing

Contributions are welcome! Please fork this repository and submit a pull request.

# PDF Extractor

This program will extract text from a PDF file and save it to a text file.
The idea is to use while developing other PDF extraction-based programs & services.

## Installation

### Using UV

After cloning the repository to your local machine,
it is recommended that UV is installed (non-uv instructions in the next section).
This makes the process much easier when working with your dependencies.
Once installed, run the following commands in order:

```bash
uv venv
source .venv/bin/activate
uv run temp.py
```

This will automatically create a virtual environment,
start the virtual environment,
and run the `temp.py` file (this action also installs any missing dependencies).

### Without UV

If UV is not able to be installed, the following commands can be used instead:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 temp.py
```

This will automatically create a virtual environment,
start the virtual environment,
install any missing dependencies,
and run the `temp.py` file.

## Usage

The application can parse the following:

- Single PDF parser.
  - Will show parsed text in the bottom window.
- Multiple PDF parser.
  - Will search all files/subfolders within the selected directory.
  - Does not show extracted text like the single PDF parser.

### Configuration

The application has configuration options to target desired values.
This essentially operates as a simple "rules engine" that can be quickly adapted.
To edit the configuration options:

- Utilize the GUI to add new rules to the configuration.
- Edit the configuration file directly.

```json
{
    "fields": [
        {
            "name": "Web Address",
            "pattern": "https?://[^\\s/$.?#].[^\\s]*",
            "required": false
        }
    ]
}
```

Considering the example above (used to extract & save web addresses):

- name: The name of the column header in the excel file.
- pattern: The regular expression (regex) used to identify the desired text.
- required: If the field is required or not.
  - If a required field is not found, the CSV will not be saved.

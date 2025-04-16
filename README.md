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

The application can has two options: Single PDF parser and Multiple PDF parser.
The multiple PDF parser will search all files/subfolders within the selected directory.
Note that the multi-PDF functionality does not currently show the extracted text;
however, the single PDF parser option will show said text in the bottom window.

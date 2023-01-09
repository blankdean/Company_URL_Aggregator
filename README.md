# Company URL Aggregator

This Python script helps you find the URLs of a list of company names by using the googlesearch python module. The script reads an Excel file containing the company names and adds the URLs to the sheet.

## Installation

To use this script, you'll need Python 3 and pip installed on your machine. You can check if you have Python 3 installed by running:

```bash
python3 --version
```

If you don't have Python 3, you can download it from the official website: https://www.python.org/downloads/

To install the required Python modules, use the provided bash script `install_dependencies.sh`:

```bash
chmod +x install_dependencies.sh
./install_dependencies.sh
```

## Usage

To run the script, use the following command:

```
python3 url_aggregator.py
```

To prevent overloading Google's servers and potentially getting your IP banned, the googlesearch module has a 2-second delay between requests. This means that the program may run somewhat slowly. You can adjust this delay by changing the `pause` parameter in the script. However, it's recommended to leave the 2-second delay in place to be on the safe side.

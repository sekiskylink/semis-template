# SEMIS Bulk Enrollment Template Generator
## Introduction
This is a python script you can use to generate a protected Excel template to use when bulk importing student enrollments into your DHIS2 SEMIS application.

## Set up
To get the script ready for use, ensure that you have the following installed
1. Python 3.6+
2. openpyxl Python library
3. dhis2.py Python library

### Installing prerequisites
```bash
pip install openpyxl
pip install dhis2.py
```
## Usage
The script `main.py` is all you need to run and generate your template. The script can be passed command line arguments that correspond to configurations you want to set before generating the template.
You can run the following to see the available options:
```bash
python main.py --help
```

This will show some documentation for each of the arguments
```text
usage: python main.py [-h] [-P PROGRAM] [-d DHIS2_URL] [-u DHIS2_USER] [-p DHIS2_PASSWORD] [-o ORGUNIT] [-s WORKBOOK_PASSWORD] [-n RECORDS]

SEMIS Bulk Enrollment Template Generator

optional arguments:
  -h, --help            show this help message and exit
  -P PROGRAM, --program PROGRAM
                        The DHIS2 tracker program
  -d DHIS2_URL, --dhis2-url DHIS2_URL
                        The DHIS2 URL with out the trailing /api
  -u DHIS2_USER, --username DHIS2_USER
                        DHIS2 username
  -p DHIS2_PASSWORD, --password DHIS2_PASSWORD
                        The DHIS2 user's password
  -o ORGUNIT, --orgunit ORGUNIT
                        The DHIS2 organisation unit ID. Comma separate multiple IDs
  -s WORKBOOK_PASSWORD, --security-password WORKBOOK_PASSWORD
                        The password to protect the workbook and its sheets. Don't use the exclamation mark '!' Defaults to semis
  -n RECORDS, --records RECORDS
                        The Number of Students

semis-template-generator
```

### Sample usage
The following two commands are the same, using the short and long arguments respectively. Use your preferred option. 
```bash
python main.py -u "admin" -p "district"  -d "https://emiseswatini.dev.hispuganda.org/emiseswatini" -P a6t4ASRXwPZ -o E63zwh4WzWK,UL2WYFdnA1p 
```

```bash
python main.py --username "admin" --password "district"  --dhis2-url "https://emiseswatini.dev.hispuganda.org/emiseswatini" --program a6t4ASRXwPZ --orgunit E63zwh4WzWK,UL2WYFdnA1p 
```



**Have fun!**
# Upwork IP Box reports generator

Checks for invoices issued on given month. For every invoice:

* If this is Fixed price - prints the notes
* If this is Hourly - downloads the time entries

Time entries are saved into CSV file, one per project

# Setup

1. [Apply for a new API key](https://www.upwork.com/services/api/apply). Required pemissions:
    * Access your basic info
    * Generate time and financial reports for your companies and teams
2. Create a local directory for detailed reports.
3. Add an alias to `.bashrc`: `alias ipbox_upwork="[ABSOLUTE_REPO_PATH]/upwork/main.py --api-key [API_KEY] --api-secret [API_SECRET] --path [REPORTS_PATH]"`
4. Copy `config.py.sample` to `config.py` and adjust the function.

# How to use?

    ipbox_upwork --month 2020-04


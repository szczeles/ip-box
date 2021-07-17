# Harvest IP Box reports scraper

Downloads Harvest time entries for a given month. Classifies entries
as IP Box applicable (or not!) and saves the detailed report to be used
durign the audit. At the ends prints sum of IP Box hours per project.

Classifier checks by default for time entry type of Development, so make sure they 
are the ones that match IP Box rules! 
You can overwrite default IP Box compatible time entry types by providing `--types` parameter.

# Setup

1. Go to [Harvest Developer Tools](https://id.getharvest.com/developers)
   and create a new personal token. Note account id and token.
2. Create a local directory for detailed reports.
3. Add an alias to `.bashrc`: `alias ipbox_harvest="[ABSOLUTE_REPO_PATH]/harvest/main.py --account [ACCOUNT_ID] --token [TOKEN] --path [REPORTS_PATH]"`

# How to use?

    ipbox_harvest --month 2020-04


# tufts-reserves-reading-list-processing-helper-scripts
This repository contains a series of helper scripts and a diagram to show how these scripts are used to help the reserves workflow to automate communication and data handoffs between departments

This repository contains a description of the Tufts workflow for receiving readings either from faculty or a university bookstore, and taking these readings (assumed to have ISBNs of some sort that can be fed into Alma Overlap Analysis) and describes how this data flows between processes and departments.   Any scripts or Analytics reports that are used for this process are prefixed and labeled either as they are here in this repo, or with the path in Analytics.

## 2ProcessOverlapAnalysis.py ##

This is the script that Tufts uses to take inventory that has been sent to us, *and processed through the overlap analysis job in Alma*  (see attached Word Doc for settings)

and takes that matching inventory and inflates it back out to include matching courses, from the request data sent in our case by the Barnes and Noble bookstore.

This use case is somewhat specific to us, but may also be adapatable to your institution's workflow if you tweak it for your data.

### Method ###

To run this script:
- assuming you have Python installed, install the neccessary requirements from the requirements file:
  - place your input file in the "Overlap Analysis Input Directory"
  - enter your institution's specific API keys in the secrets_local.py file in the config directory
  - ```python3 -m pip install -r requirements.txt```
    - or ```python pip install -r requirements.txt```
  - ```python3 2ProcessOverlapAnalysis.py```
### Output ###
  - in the "Barnes and Noble Parsed" directory


Note also that to ease the heavy lifting of adding citations to reading lists, Tufts also developed the Citation Loader cloud app to take a spreadsheet of MMS IDs and course codes (can also have barcodes for physical items, and you can specify a temporary reserves location to move these items to, as well as citation type and reading list section).  This is also now integral to our process and is referenced here in this diagram



![Workflow Diagram](https://www.library.tufts.edu/reserves_workflows/Reserves_Workflows_at_Tufts.svg)

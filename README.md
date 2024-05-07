# tufts-reserves-reading-list-processing-helper-scripts
This repository contains a series of helper scripts and a diagram to show how these scripts are used to help the reserves workflow to automate communication and data handoffs between departments

This repository contains a description of the Tufts workflow for receiving readings either from faculty or a university bookstore, and taking these readings (assumed to have ISBNs of some sort that can be fed into Alma Overlap Analysis) and describes how this data flows between processes and departments.   Any scripts or Analytics reports that are used for this process are prefixed and labeled either as they are here in this repo, or with the path in Analytics.

Note that in a v2 of this process, we intend to be able to take citations that don't have ISBN, say from email, and incorporate Micrsoft Copilot to parse them, whatever citation style they are, into bibliographic metadata that then uses a website that calls the Primo API to get the data that can be fed into this process.

Note also that to ease the heavy lifting of adding citations to reading lists, Tufts also developed the Citation Loader cloud app to take a spreadsheet of MMS IDs and course codes (can also have barcodes for physical items, and you can specify a temporary reserves location to move these items to, as well as citation type and reading list section).  This is also now integral to our process and is referenced here in this diagram

![workflow diagram](https://www.library.tufts.edu/reserves_workflows/Reserves Workflows at Tufts.svg)

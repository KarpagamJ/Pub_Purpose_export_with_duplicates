This script is designed to fetch all "Purposes" from the OneTrust Consent Manager API (with pagination), normalize their labels, check for duplicates, and then export everything to an Excel file with two sheets:

All Purposes → contains all purposes fetched from the API.
Label Duplicates → contains only potential duplicates based on label similarity (ignoring spacing differences).

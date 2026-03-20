# Sync Link Worksets

## Description
This tool synchronizes the host model worksets with the worksets detected in loaded Revit links.

It can:

- Create missing worksets from linked models  
- List worksets present in links  
- Detect unused host worksets  
- Delete empty unused host worksets when allowed by Revit  

The tool does **not modify existing worksets**. It only reports, creates, or deletes worksets depending on the selected mode.

---

## Modes

### RUN: CREATE MISSING WORKSETS
Creates host worksets that exist in the linked models but are missing in the host model.

### PROBE: LIST LINK WORKSETS
Lists all worksets detected in loaded Revit links and indicates whether they already exist in the host model.

### PROBE: LIST UNUSED HOST WORKSETS
Scans host worksets and reports which ones are unused or empty.

### RUN: DELETE UNUSED HOST WORKSETS
Deletes empty host worksets when Revit allows deletion.  
Protected worksets or worksets with elements will not be deleted.

---

## Usage

1. Click the **Sync Link Worksets** button.
2. Choose the desired mode.
3. The tool scans the model and linked files.
4. A detailed report is printed in the **pyRevit output window**.

---

## Notes

- The host model must be **workshared**.
- Only **User Worksets** are processed.
- Protected or default worksets are ignored.
- Worksets containing elements will not be deleted.

---

## Author

Ricardo J. Mendieta  
pyMENVIC
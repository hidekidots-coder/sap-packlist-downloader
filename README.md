# SAP Packing List Downloader

VBA automation that downloads **Packing List** attachments from SAP (transaction **VT03N**) based on a list of Transport Numbers in Excel.

## Overview

This project automates the process of opening transport documents in SAP, searching for Packing List attachments and saving them as PDF files.

## Problem

Manually opening each transport in VT03N, navigating to the attachment list and downloading the Packing List is repetitive and time-consuming.

## Solution

The macro:

- Reads Transport Numbers from an Excel sheet
- Opens each document in transaction VT03N
- Searches for attachments containing "PACK" or "LIST" in the title
- Downloads the file as PDF
- Marks the result as **OK** or **ERRO** in column B

## Workflow
Excel Sheet (PACKINGLIST)
↓
Read Transport Number (TKNUM)
↓
SAP Transaction VT03N
↓
Open Attachment List (GOS)
↓
Find Packing List
↓
Download PDF
↓
Mark OK / ERRO

## Requirements

- SAP GUI with Scripting enabled
- User already logged into SAP
- Excel sheet named `PACKINGLIST`
- Column A containing the Transport Numbers (starting from A2)

## How to Use

1. Open the Excel file
2. Rename the sheet to `PACKINGLIST` (if necessary)
3. Place the Transport Numbers in column A (starting at A2)
4. Press `Alt + F11` and import `src/Loop_SAP_Download_Packlist.bas`
5. Run the macro `Loop_SAP_Download_Packlist`
6. The PDFs will be saved in:  
   `C:\Users\<seu_usuario>\Downloads\Intercoout\`

## Output

| Column | Content                  |
|--------|--------------------------|
| A      | Transport Number (TKNUM) |
| B      | Result (`OK` or `ERRO`)  |

## Notes

- The macro automatically creates the download folder if it does not exist
- If a file with the same name already exists, a timestamp is added to avoid overwriting
- Keep the SAP window visible while the macro is running

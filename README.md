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

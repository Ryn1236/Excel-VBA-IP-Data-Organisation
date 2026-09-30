## Excel VBA Scan Data Automation

## About

This project uses Excel VBA to process and organise scan data based on IP addresses. It was created during my internship to make it easier to handle large amounts of data and reduce manual sorting.

The tool can process Excel or CSV files and automatically create separate worksheets for each IP address. Rows without a valid IP address are placed in a `No IP` worksheet.

## Technologies

* Microsoft Excel
* VBA
* Regular Expressions

## What It Does

* Allows the user to select an Excel or CSV file
* Checks Column D for valid IPv4 addresses
* Creates a new worksheet for each unique IP address
* Sorts the related rows into the correct worksheet
* Places rows without a valid IP address into a `No IP` worksheet
* Processes multiple worksheets from the selected file
* Opens the original file as read-only

## How It Works

1. Run the `InsertFile` macro.
2. Select the Excel or CSV file to process.
3. The macro reads the data and checks Column D for IP addresses.
4. A worksheet is created for each unique IP address.
5. The related rows are copied into their respective worksheets.
6. Rows without a valid IP address are placed in the `No IP` worksheet.
7. The original file is closed without making changes.

## Example

The tool can identify and organise IP addresses such as:

```text
192.168.1.1
10.0.0.25
172.16.0.10
```

Invalid or missing IP addresses are placed in the `No IP` worksheet.

## Project Outcome

The tool was used to help process large amounts of scan data, including datasets with around **400,000 rows**, making the sorting and organisation process faster and less manual.

## Disclaimer

This repository contains a demonstration version of the project. Any company, government, or confidential information has been removed or replaced with sample data.


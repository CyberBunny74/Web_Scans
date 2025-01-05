# Web Server Port Scanner #
This repository contains two PowerShell scripts for scanning web server ports across different network scopes. These scripts are useful for network administrators and IT professionals who need to perform port checks on web servers.

## Scripts ##
  ### Web.ps1 ### 
This script performs port checks on a list of specified computers for HTTP (80) and HTTPS (443) ports.

Features:
* Imports computer list from a CSV file
* Checks ports 80 and 443 for each computer
* Exports results to an Excel file

Usage:
* Replace %INPUT_CSV% with the path to your input CSV file containing computer names.
* Replace %OUTPUT_EXCEL% with the desired path for the output Excel file.
* Run the script in PowerShell.

  ### Web_by_subnet.ps1 ###
This script conducts a comprehensive port scan on an entire subnet, focusing on web server ports.

Features:
* Scans a specified subnet range
* Configurable port selection
* Error logging and progress tracking
* Hostname resolution for IPs with open ports
* Exports results to an Excel file

Usage:
* Replace %SUBNET% with your target subnet (e.g., "192.168.1.0/24").
* Set %WEB_PORTS% to the ports you wish to check (e.g., "80, 443").
* Configure the following file paths:
  * %ERROR_LOG%: Path for the error log file
  * %PROGRESS_JSON%: Path for the progress tracking file
  * %RESULTS_EXCEL%: Path for the output Excel file
* Run the script in PowerShell.

Requirements
* PowerShell 5.1 or later
* ImportExcel module (install via Install-Module ImportExcel)

Installation
* Clone this repository or download the scripts.
* Ensure you have the required PowerShell version and ImportExcel module installed.
* Modify the placeholder variables in the scripts as needed.


## Note ##
These scripts are provided as-is and should be used responsibly and in compliance with your organization's policies and applicable laws.

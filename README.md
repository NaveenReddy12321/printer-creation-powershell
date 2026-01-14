\# Bulk Printer Creation using PowerShell



\## Overview

This project provides a PowerShell automation script to bulk create printers on a Windows Print Server using a CSV input file.



The script is designed for Windows system administrators to eliminate manual printer creation, reduce configuration errors, and standardize printer deployments.



---



\## Features

\- Bulk printer creation from CSV

\- Automatic TCP/IP printer port creation

\- Assigns printers to correct ports

\- Supports printer sharing configuration

\- Prevents duplicate printer creation

\- Clear console output for success and failures

\- Generates CSV logs for every execution



---



\## Prerequisites

\- Windows Server with Print Management role

\- PowerShell 5.1 or later

\- Script must be run as Administrator

\- Printer drivers must already be installed

\- Print Spooler service must be running



---



\## CSV Format

The script expects the following columns:



| Column Name | Description |

|-----------|-------------|

| PrinterName | Name of the printer |

| PrinterIP | IP address of the printer |

| PortName | TCP/IP port name |

| Comment | Printer comment |

| Description | Printer description |

| Shared | Yes / No |

| DriverName | Installed printer driver name |



Example:

```csv

PrinterName,PrinterIP,PortName,Comment,Description,Shared,DriverName

HP-Floor1,192.168.1.20,IP\_192.168.1.20,Floor 1 Printer,HP LaserJet Floor 1,Yes,HP Universal Printing PCL 6



\## How to Use



* Copy the CSV file to the server
* ```

Example:

C:\\Temp\\printers.csv```



* Open PowerShell as Administrator

```

Allow script execution:

Set-ExecutionPolicy RemoteSigned -Scope Process```



* Run the script:

```

.\\Create-Printers.ps1```



\## Logging



Each execution generates a CSV log file:

```

C:\\PrinterLogs\\PrinterCreation\_YYYYMMDD\_HHMMSS.csv```





* The log contains:
* Timestamp
* Printer name
* Printer IP
* Port name
* Status (Success / Failed / Skipped)
* Error message (if any)



\## Use Cases



* New office printer onboarding
* Print server rebuilds
* Bulk printer migrations
* Standardized printer deployments




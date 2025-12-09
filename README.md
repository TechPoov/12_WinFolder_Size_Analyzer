WinFolderSizeAnalyzer Overview
==============================

WinFolder_Size_Analyzer is a lightweight TechPoov utility that scans Windows folders and generates clean CSV reports with aggregated folder sizes and file counts.  

It supports Files-only, Folders-only, or Both modes, and allows MaxDepth control for targeted scanning.

The tool produces three outputs per job --- a **Data CSV**, a **Log CSV**, and an optional **Debug Log** --- along with a run-level **Summary CSV**.  

Designed for Windows Task Scheduler, it runs silently without pop-ups, making it ideal for automated audits, migrations, and storage analysis.

Links:
------

-   [Download Tool](https://github.com/TechPoov/12_WinFolder_Size_Analyzer/releases/) / ![Downloaded Count](https://img.shields.io/github/downloads/TechPoov/12_WinFolder_Size_Analyzer/total.svg)

-   [View User Manual](https://docs.google.com/document/d/1gZcs_K-3kZaK6Epvx8G-e-jJJKzC0wJCt16G7aWlUb8/edit?tab=t.0) / [View Tool Comparison](https://docs.google.com/document/d/1IDzA_8dCJ7YXy37MtVZAy9ZXve4DXtUl7rZnH1H7eDg/edit?tab=t.0) / [View Test Cases](https://docs.google.com/spreadsheets/d/1Km3BTzkHM5ycnRgYB_aL7IXAB1inN4QABwWl6JcPt0o/edit?gid=0#gid=0)

Features
========

Key Features (Google Docs Version)
==================================

Folder Size Aggregation\
Recursively computes full folder sizes, including all nested subfolders and files.

File & Folder Scanning Modes\
Supports three modes: FILES, FOLDER, and BOTH, allowing flexible inventory options.

MaxDepth Control\
Allows limiting recursion depth to scan only specific directory levels.

Multi-Job Execution\
Runs multiple independent scan jobs from a single .config file.

Clean CSV Outputs\
Generates structured Data CSV files containing paths, timestamps, sizes, and file counts.

Detailed Log CSV\
Each job creates a timestamped Log CSV capturing status updates, warnings, errors, and execution steps.

Optional Debug Log\
When enabled, a detailed debug file logs internal steps for troubleshooting.

Run-Level Summary File\
Produces a summary CSV listing all jobs, their status, durations, and output paths.

Zero Installation Required\
Runs natively on Windows using wscript or cscript with no admin rights or installations.

Task Scheduler Friendly\
Executes silently without pop-ups---ideal for automation, recurring audits, and background scanning.

Supports Local and Network Paths\
Works with local drives, mapped drives, and UNC paths such as \\server\share.

Robust Error Handling\
All errors are captured in the log file; the tool continues running safely without breaking the job.

Who Needs This Tool & How They Use It
=====================================

1\. Storage Administrators Who Need Accurate Space Analysis
-----------------------------------------------------------

Administrators responsible for servers, shared drives, and storage infrastructure need a fast, reliable way to measure how space is being used across deep folder trees. This tool provides aggregated sizes and file counts without installing GUI software.

Use Cases:\
- Identifying which folders consume the most disk space.\
- Preparing reports for storage expansion or cleanup projects.\
- Pinpointing high-growth directories before they trigger critical alerts.\
- Understanding size distribution before reorganizing large data sets.

* * * * *

2\. Migration & Transition Teams
--------------------------------

Teams planning to move data to new servers, cloud platforms, or archival drives must first understand the source folder structure and its total size. This tool provides clean, predictable CSVs essential for estimating effort, cost, and timelines.

Use Cases:\
- Measuring total volume to migrate from local storage to cloud.\
- Identifying deep or complex folder structures requiring special migration handling.\
- Creating pre-migration inventory reports for stakeholders.\
- Validating that the data footprint matches post-migration results.

* * * * *

3\. IT Support & Operations Teams
---------------------------------

Helpdesk and operations teams often get requests related to "disk full," "slow drive," or "mysterious large folders." This tool becomes a diagnostic companion for quick, structured analysis.

Use Cases:\
- Finding unusually large folders when troubleshooting performance issues.\
- Generating logs to investigate user complaints about storage usage.\
- Running scheduled size audits to prevent unexpected low-space incidents.\
- Creating evidence for IT governance or internal review processes.

* * * * *

4\. Audit, Compliance, and Documentation Teams
----------------------------------------------

These teams need timestamped, structured documentation of directory contents for regulatory requirements, internal controls, or annual audits.

Use Cases:\
- Producing inventory snapshots of shared or departmental data.\
- Creating audit trails that show structure, timestamps, and hierarchical content.\
- Comparing data footprint year-to-year for compliance reviews.\
- Generating clean CSVs that can be attached to audit reports.

* * * * *

5\. Project Managers & Documentation Builders
---------------------------------------------

PMs handling large shared folders (project archives, client deliverables, training materials) need clarity on what data exists and how big it is. This tool provides quick visibility without manually checking each folder.

Use Cases:\
- Documenting folder structure for project handovers.\
- Reviewing old project archives to decide what to retain or delete.\
- Preparing team-wide data snapshots during major transitions.\
- Understanding folder growth across the project lifecycle.

* * * * *

6\. Power Users & Technical End Users
-------------------------------------

Users who maintain large personal or work-related folder structures benefit from a simple, portable way to analyze their own data footprint.

Use Cases:\
- Tracking personal archive sizes across years.\
- Inspecting large collections (photos, videos, documents) before cleanup.\
- Comparing multiple drives or partitions for reorganization.\
- Generating CSVs to visualize size distribution in Excel or BI tools.

* * * * *

7\. Administrators in Restricted Server Environments
----------------------------------------------------

Some environments do not allow installing TreeSize, WinDirStat, or other tools. Because WinFolder_Size_Analyzer runs from a plain .vbs file, it works even on locked-down enterprise servers.

Use Cases:\
- Running audits on terminal servers where installation is not allowed.\
- Generating reports on cloud-hosted VMs without GUI access.\
- Using Task Scheduler to perform daily/weekly silent scans.\
- Creating automated CSV outputs for SIEM or monitoring pipelines.

Requirements
============

* * * * *

1\. Functional Requirements
---------------------------

Folder Scanning\
The tool must scan a given root folder and all subfolders recursively.

Modes Support\
Must support three scan modes: FILES, FOLDER, BOTH.

Folder Size Aggregation\
The tool must calculate total folder size, including all nested content.

File Count Tracking\
The tool must count all files inside each folder and write the value to the CSV.

MaxDepth Control\
The tool must respect the configured maximum recursion depth (0 = unlimited).

Multi-Job Execution\
The tool must execute multiple scan jobs defined in a single .config file.

Output Generation\
Must produce a Data CSV, a Log CSV, and optionally a Debug log for each job.

Run-Level Summary\
Must produce a summary CSV containing status, start/end times, counts, and output paths for all jobs.

Graceful Error Handling\
The tool must continue scanning even if some folders cannot be accessed.

Silent Operation\
There must be no pop-ups or prompts during execution to support scheduled runs.

* * * * *

2\. Non-Functional Requirements
-------------------------------

Performance\
The tool should complete scans efficiently even on large directory structures.

Reliability\
Must produce consistent and repeatable results across runs.

Error Resilience\
Should log all failures without stopping the execution of other jobs.

Portability\
Must run on any Windows system without installation or external libraries.

Security\
Should respect existing file system permissions and must not modify data.

Resource Efficiency\
Should use minimal memory and avoid loading entire folder trees at once.

Automation Compatibility\
Must run safely through Windows Task Scheduler and similar tools.

* * * * *

3\. User Experience Requirements
--------------------------------

Simple Configuration\
Users should configure all scan jobs via a plain-text .config file.

Readable Output\
CSV files must use consistent column ordering and clear naming.

Accessible Logs\
Logs should be easy to interpret for troubleshooting.

Predictable File Naming\
Output filenames must include timestamp and job name for easy identification.

Unattended Execution\
The tool should run to completion without user interaction.

* * * * *

4\. System Requirements
-----------------------

Operating System\
Windows 7, Windows 10, Windows 11, Windows Server environments.

Runtime\
VBScript enabled (wscript.exe / cscript.exe).

Permissions\
Read access to the scan folder.\
Write access to the output folder and script directory.

Environment\
Support for local disks, mapped drives, and UNC paths.

No Installation Needed\
The script must operate as a standalone .vbs file.

Get Tool and Documents
=============================

-   [Tool](https://docs.google.com/spreadsheets/d/1fVHxhQlCNSWalqE9CUgprPba4uBarCpNdC5XlMKR_VA/copy)

-   [User Manual V1.0](https://docs.google.com/document/d/1PdNsuZ_GIPac5KTKAA4F6KEvlfco9pCQfAsLotnuu9A/copy)

-   [Test Cases for reference](https://docs.google.com/spreadsheets/d/1nKbSnOrZPNFgR2kuCPCXXKIjS4Tm8pd8fgmh6u2J1Ew/copy)

Version History
===============

V1.0 --- 01- Dec - 2025   Initial release

* * * * *

License
=======

This project is released under the MIT License, a widely used open-source license that allows personal, commercial, and organizational use with minimal restrictions. Users are free to use, modify, distribute, and incorporate the code into their own projects, provided that the original copyright notice and license terms are included in all copies or substantial portions of the software.

* * * * *

Support
=======

If you need help, have questions, or want to share feedback, support is always available.

🌐 Online Documents:\
Visit the official TechPoov documentation at https://techpoov.github.io

📧 Email Support:\
techpoov+GDrive-FolderCopy@gmail.com

🐞 Report Issues on GitHub:\
Submit bugs, feature requests, or enhancement ideas here:\
https://github.com/TechPoov/GDrive-FolderCopy/issues

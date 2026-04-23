# IP-Intel
A bulk IP address lookup tool that quickly gathers information from ARIN, the American Registry for Internet Numbers for investigative purposes

## Intended Use
Have you ever found yourself looking at a large IP address dataset, attempting to find any actionalbe intelligence?  Instead of manually checking hundreds of IP addresses for registry and location information, let IP Intel handle the heavy lifting for you.  Results output to an easy to read spreadsheet and an interactive web browser map.  

## Screenshots

<p align="center">
  <img src="Screenshots/Screenshot_2026-04-23_100539.png" width="19%" />
  <img src="Screenshots/Screenshot_2026-04-23_100723.png" width="19%" />
  <img src="Screenshots/Screenshot_2026-04-23_100805.png" width="19%" />
  <img src="Screenshots/Screenshot_2026-04-23_100817.png" width="19%" />
  <img src="Screenshots/Screenshot_2026-04-23_100919.png" width="19%" />
</p>

## Overview
This script is a desktop application for processing lists of IP addresses. It validates IPs, enriches them with geo and ARIN data, and then generates a map and a final report.

IP Intel lets the user download an evidence template in Excel format and enter their data manually.  The user can also select a source file in .csv, .xls, .xlsx, or .ods format, although not all formats are compatible.  IP Intel will then validate IPs using Python’s ipaddress module and will remove duplicates so each IP is processed once.  It then queries external services for each IP to collect location and organization details.  IP Intel exports results to .xlsx and an interactive HTML map when latitude / longitude are available.

## Workflow
1. Input selection
The user chooses a file containing IP-related evidence. The program enables the processing button only after a file is selected.

2. Validation
It strips whitespace, filters invalid IPs, and drops duplicates before processing begins.

3. Enrichment
Each IP is sent to IP-API for geolocation and to ARIN for organization lookup.

4. Output
The app saves a spreadsheet report and, if geolocation information is present, a map HTML file in the chosen output folder.

## Program Dependencies
os, sys, threading, datetime, ipaddress	Standard-library support for filesystem access, packaging checks, background processing, timestamps, and IP valida

## Under the hood
IP Intel is simply a python script with a GUI for ease of use, packaged in a portable executable file for Windows.  IP Intel requires NO installation and can be ran directly on the host machine or from an external drive / USB.  

##  Open Source Software 
This program is free and open source, and it may be modified, shared, and used freely.

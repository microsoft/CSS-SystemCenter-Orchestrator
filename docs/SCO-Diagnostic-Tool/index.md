# SCO Diagnostic Tool

Download the latest release:  [SCO-Diagnostic-Tool.ps1]({{ site.latestDownloadLink }}/SCO-Diagnostic-Tool.ps1)

## Description

The *SCO Diagnostic Tool* allows you to collect diagnostic logs from your Orchestrator environment to help you and Microsoft technical support engineers to resolve Orchestrator technical incidents faster. It is a light, script-based, open-source tool.

# How to run

1. **Log on** to the Orchestrator mgmt. server with an admin account, preferably with the service account of the "Orchestrator Management Service".
2. **Save** the script into a folder.
3. Execute the script with right/click + **"Run with PowerShell"**.
   > ##### Note: 
   > If PowerShell starts and quits immediately, then you need to run the following command in a RunAsAdmin PowerShell window:
   ````
   Set-ExecutionPolicy RemoteSigned
   ````
4. Follow the instructions.
  > The script may ask to manually input the Orchestrator SQL Server instance/database if not able to automatically detect it.
5. **Upload** the resulting zip file to Microsoft CSS.

## Minimum requirements

- Windows Server 2016 or later
- Orchestrator 2019 or later
- Windows Powershell version 4.0 or later

## Notes:

- The script won't make any change in the SCSM environment.
- The script can be also used as a Health Checker. 

## Do you want to contribute to this tool?

[Here]({{ site.GitHubRepoLink }}/SCO-Diagnostic-Tool) is the GitHub repo.

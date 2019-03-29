# Sharepoint CSOM DLL's

### This folder contains all the libraries that are necessary for the scripts works.

These are only used in case of not have installed the [SharePoint Online Client Components SDK](https://www.microsoft.com/en-us/download/details.aspx?id=42038) in the computer that will execute the scripts.

**NOTE**: By default, all the scripts has commented the portion of code that loads the libraries on runtime, inside the *Global Script Configuration* region


##### Code Block
```
-- Files Path Configuration
$currentScriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$parentCurrentPath = split-path -parent $currentScriptPath
$currentScriptPathParent = split-path -parent $currentScriptPath
$folderBase = $currentScriptPathParent + "\SharePoint DLL\"
$logFileCreated = Get-Date -format 'yyyy_MM_dd_hhmm' 
$Logfile = $currentScriptPath + '\' + $logFileCreated + "_log"

-- Adding the Client Object Model Assemblies   
$sCSOMRuntimePath=$folderBase +  "Microsoft.SharePoint.Client.Runtime.dll"         
$sCSOMPath=$folderBase +  "Microsoft.SharePoint.Client.dll"
$sCSOMPath2=$folderBase +  "Microsoft.SharePoint.Client.WorkflowServices.dll"
Add-Type -Path $sCSOMPath          
Add-Type -Path $sCSOMPath2 
Add-Type -Path $sCSOMRuntimePath   
```

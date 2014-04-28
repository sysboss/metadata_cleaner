MS Office Metadata Cleaner
====

## Description ##
Removes personal credentials (document author, company name etc.) from Microsoft Office files.  
Automatic files scan in all system drives. Designed to run as scheduled task.  

Cleans MS Office registry keys:  
  hkcu:Software\Microsoft\Office\Common\UserInfo\Company  
  hkcu:Software\Microsoft\Office\Common\UserInfo\UserInitials  
  hkcu:Software\Microsoft\Office\Common\UserInfo\UserName  
  
Supports *.xls, *.xlsx, *.doc, *.docx, *.pptx, *.ppt, *.dot, *.dotx  

## Usage ##
Options:

```
  -help|h          Display help
  -verbos          Verbose mode
  -backup PATH     Backup all original files to PATH before any changes
  -dryrun          Do not save changes (demonstration mode)
```

To make it run:
> Set-ExecutionPolicy Unrestricted  

Answer 'Y' or just press ENTER.
  
If used as scheduled task run as shown:
```
 Name:     MetaData-Cleaner
 Triggers: Daily
 Actions:
    Start a program: powershell.exe
    Arguments:       -WindowStyle "Hidden" -noprofile -executionpolicy bypass -file C:\office_metadata_cleaner.ps1
```

## System Requirements ##
PowerShell v2 or above  

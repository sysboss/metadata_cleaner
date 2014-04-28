# -------------------------------------------------
# MS Office Meta data cleaner
# Copyright (c) 2014 Alexey Baikov <sysboss@mail.ru>
#
# Remove Personal Document Information
# -------------------------------------------------
# Version: 0.2
#
# Make it run
# Set-ExecutionPolicy Unrestricted
# powershell.exe -noprofile -executionpolicy bypass -file C:\Windows\metadatacleaner\office_metadata_cleaner.ps1

# Parameters
Param(
    [switch]$help,
    [switch]$usage,
    [switch]$h,
    
    [Parameter(Mandatory=$False)]
    [string]$backup,
	
    [switch]$dryrun
)

$Products = 3;
cls

# Initialize log file
Try{ Stop-Transcript | out-null } Catch {}

if( !(Test-Path -Path 'C:\Logs') ){
	New-Item -ItemType directory -Path 'C:\Logs'
}

Start-Transcript -path "C:\Logs\msoffice-$(gc env:computername)-cleaner.log" | out-null


# Window style
$a = (Get-Host).UI.RawUI
$a.ForegroundColor = "white"
$a.WindowTitle     = "Personal Information Cleaner"

$PSv = $psversiontable.PSVersion.Major
$USR = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

"
# MS Office
# Remove Personal Document Information
#
# PowerShell v.$PSv (Current user: $USR)

"

if($help -or $usage -or $h){
    '
usage: '+$MyInvocation.InvocationName+' [options] FROM
  -help|h          Help (this info)
  -verbos          Verbose mode
  -backup PATH     Backup all original files to PATH before
  -dryrun          Do not save changes (demonstration mode)
    '
    exit 1
}

# Set registry
set-itemproperty -ErrorAction SilentlyContinue -Path hkcu:Software\Microsoft\Office\Common\UserInfo -Name "Company" -value ""
set-itemproperty -ErrorAction SilentlyContinue -Path hkcu:Software\Microsoft\Office\Common\UserInfo -Name "UserInitials" -value ""
set-itemproperty -ErrorAction SilentlyContinue -Path hkcu:Software\Microsoft\Office\Common\UserInfo -Name "UserName" -value ""

# Excel Object
Try {
	Add-Type -AssemblyName Microsoft.Office.Interop.Excel
	$xlRemoveDocType = "Microsoft.Office.Interop.Excel.XlRemoveDocInfoType" -as [type]
	$objExcel        = New-Object -ComObject excel.application

	Try { $objExcel.visible       = $false
	      $objExcel.DisplayAlerts = false
    } Catch {}
} Catch {
	"[*] Excel is not installed."
	Write-Verbose $_.Exception.Message
	$Products--;
}

# Word Object
Try {
	Add-Type -AssemblyName Microsoft.Office.Interop.Word
	$RemoveDocType = "Microsoft.Office.Interop.Word.WdRemoveDocInfoType" -as [type]
	$objWord       = New-Object -ComObject Word.Application

	Try { $objWord.visible       = $false
	      $objWord.DisplayAlerts = false
    } Catch {}
} Catch {
	"[*] Word is not installed."
	Write-Verbose $_.Exception.Message
	$Products--;
}

# PowerPoint Object
Try {
	Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
	$RemoveDocTypePp     = "Microsoft.Office.Interop.PowerPoint.PpRemoveDocInfoType" -as [type]
	$objPp               = New-Object -ComObject PowerPoint.Application
	Try { $objPp.Visible = $false } Catch {}
} Catch {
	"[*] PowerPoint is not installed."
	Write-Verbose $_.Exception.Message
	$Products--;
}

# if non of object exist
if ($Products -le "0"){
	"
	None of supported products in installed. Please check you Office installation.
	Abort.
	"
	exit 0;
}

if( $backup ){
    if( !(Test-Path $backup) ){
        Write-Verbose "$backup no such directory. Creating..."
        mkdir $backup | out-null

        if( !(Test-Path $backup) ){
            "Failed to create $backup folder. Exit."
            Stop-Transcript | out-null
            exit 2
        }
    }
}

# Errors Handling
$ErrorActionPreference= 'silentlycontinue'

# Functions
function Test-IsWritable(){
	[CmdletBinding()]
	param([Parameter(Mandatory=$true,ValueFromPipeline=$true)][psobject]$path)
	
	process{
		Write-Verbose "Test if file $path is writeable"
		if (Test-Path -Path $path -PathType Leaf){
			Write-Verbose "File is present"
			$target = Get-Item $path -Force
			Write-Verbose "File is readable"
			try{
				Write-Verbose "Trying to openwrite"	
				$writestream = $target.Openwrite()
				Write-Verbose "Openwrite succeded"	
				$writestream.Close() | Out-Null
				Write-Verbose "Closing file"				
				Remove-Variable -Name writestream
				Write-Verbose "File is writable"
				Write-Output $true
				}
			catch{
				Write-Verbose "Openwrite failed"
				Write-Verbose "File is not writable"
				Write-Output $false
				}
			Write-Verbose "Tidying up"
			Remove-Variable -Name target
		}
		else{
			Write-Verbose "File $path does not exist or is a directory"
			Write-Output $false
		}
	}
}

$drives  = Get-PSDrive -PSProvider filesystem
$items   = @()

Write-Verbose "" + $drives.count + " Drives detected ($drives)"

"Seeking files... please be patient."

foreach ($DriveLetter in $drives) {
    $drive = "" + $DriveLetter + ':\'
    $result = Get-Childitem $drive -Recurse -include *.xls, *.xlsx, *.doc, *.docx, *.pptx, *.ppt, *.dot, *.dotx
    if ($result){
        $items += $result;
    }
}

$ExcelProcess = get-process excel
$WordProcess  = get-process word
$PPProcess    = get-process powerpnt

$PPProcess.StartInfo.WindowStyle="Hidden"

$count = $items.count
$i     = $count

"" + $count + " file(s) found."

foreach ($item in $items) {
    if (($item.FullName.Contains('Microsoft Office'))){
        Write-Verbose "$logstring" + $item.FullName + " is built-in Office file. Skip."
        $i--
        continue
    }

    $Time      = Get-Date
    $logstring = "$Time ("+($count-$i+1)+"/$count) `t"
    $isOK      = Test-IsWritable -path $item.FullName
    
    if ($item.isreadonly){
        "$logstring" + $item.FullName + " is ReadOnly file. Cannot be edited."
        $i--;
        continue;
    }
    if(!$isOK){
        "$logstring " + $item.FullName + " can't be modified by $USR. Skip."
        $i--;
        continue;
    }else{
        if ($backup){
            $path = $backup+"\"+$item.name
            if (Test-Path $path){
                $path = $backup+"\"+$i+"_"+$item.name
                Copy-Item -Path $item.FullName -Destination $path
            }else{
                Copy-Item -Path $item.FullName -Destination $path
            }
        }
		
        if (($item.Extension -like '.xls') -or ($item.Extension -like '.xlsx')) {
            Try 
            {
				$objExcel.DisplayAlerts = $false
                $workbook = $objExcel.workbooks.open($item.fullname)
				$workbook.Password = ""
                "$logstring EXECL: $item"
				
                $objExcel.visible = $false
                $workbook.RemoveDocumentInformation($xlRemoveDocType::xlRDIAll)
                if(!$dryrun){ $workbook.Save() }
                $objExcel.Workbooks.close();
            }
            Catch
            {
                $ErrorMessage = $_.Exception.Message
                if ($ErrorMessage.Contains('read-only')){
                    "$logstring ERROR file is read-only or old office format $item"
                }elseif($ErrorMessage.Contains('The password is incorrect')){
					"$logstring ERROR file is password protected $item"
				}else{
                    "$logstring ERROR Failed to EXCEL RemoveDocType from: $ErrorMessage"
                }
				$objExcel.Workbooks.close();
            }
        }
        elseif (($item.Extension -like '.doc') -or ($item.Extension -like '.docx')) {
            Try 
            {
                $wordbook = $objWord.Documents.Open($item.fullname,$false,$false,$false,"***","***")
				$wordbook.Password = ""
                "$logstring WORD: $item"
				
                $objWord.visible  = $false
                $wordbook.RemoveDocumentInformation($RemoveDocType::wdRDIAll)
                if(!$dryrun){ $wordbook.Save() }
                $objWord.ActiveDocument.Close()
            }
            Catch
            {
                $ErrorMessage = $_.Exception.Message
                if ($ErrorMessage.Contains('read-only')){
                    "$logstring ERROR file is read-only or old office format $item"
                }elseif($ErrorMessage.Contains('The password is incorrect')){
					"$logstring ERROR file is password protected $item"
				}else{
                    "$logstring ERROR Failed to WORD RemoveDocType from: $ErrorMessage"
                }
				$objWord.ActiveDocument.Close()
            }
        }
        elseif (($item.Extension -like '.ppt') -or ($item.Extension -like '.pptx')) {
            Try 
            {
                $ppbook = $objPp.Presentations.Open($item.fullname)
                $ppPres = $objPp.ActivePresentation
                "$logstring PowerPoint: $item"
                $ppPres.RemoveDocumentInformation($RemoveDocTypePp::ppRDIAll)
                if(!$dryrun){ $ppbook.Save() }
                $ppPres.Close()
            }
            Catch
            {
                $ErrorMessage = $_.Exception.Message
                if ($ErrorMessage.Contains('read-only')){
                    "$logstring ERROR file is read-only or old office format $item"
                }else{
                    "$logstring ERROR Failed to PowerPoint RemoveDocType from: $ErrorMessage"
                }
            }
        }
        else {
            "$logstring ERROR File " + $item.fullname + " is unsupported here."
        }
    }
    $i--;
}

$objExcel.Quit() | out-null
$objWord.Quit()  | out-null
$objPp.Quit()    | out-null

if ($PPProcess)   { $PPProcess.Kill()    }
if ($ExcelProcess){ $ExcelProcess.Kill() }
if ($WordProcess) { $WordProcess.Kill()  }

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | out-null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWord)  | out-null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objPp)    | out-null

[gc]::collect() | out-null
[gc]::WaitForPendingFinalizers() | out-null

exit

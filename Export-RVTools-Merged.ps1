<#
Export-RVTools.ps1 - Run RVTools with multiple vCenter and merge exported files.
v0.1 2021/6/20

- Requirements:
PowerShell
RVTools - For XLSX export and merge

- Optional:
SharePoint PnP - For uploading to SharePoint

- vcenter.csv format:
1: VCenter,User,EncryptedPW
2: vcenter1,user1,encryptedpw1
...

Note: Generate encrypted password with RVToolsPasswordEncryption.exe

- Modify variables before first run
$RvtoolsDir - RVTools install directory
$ConfigDir - vcenter.csv directory
$OutputDir - Output directory

#>

$RvtoolsDir = "D:\Apps\RVTools\Robware\RVTools"
$RvtoolsExe = "${RvtoolsDir}\RVTools.exe"
$RvtoolsMergeExe = "${RvtoolsDir}\RVToolsMergeExcelFiles.exe"

$ConfigDir = "D:\Work\Scripting\Export-RVTools"
$OutputDir = "D:\Work\__Scheduled__rvtools_export"

$DateStr = Get-Date -Format "yyyyMMdd_HHmmss"
$MergedOutputFile = "RVTools_AllVC_$DateStr.xlsx"

$SharePointURL = ""
$SharePointPath = ""

Set-Location $ConfigDir
$Csvfile = Import-Csv ".\vcenter.csv"
$Filelist=@()

# Export for each vCenter
Set-Location $OutputDir
$Csvfile | % {

	$Server = $_.VCenter
	$User = $_.User
	$Pw = $_.EncryptedPW

	$Filename = "RVTools_${Server}_${DateStr}.xlsx"

	$FileList += $Filename
	$ArgumentList = "-s $Server -u $User -p $Pw -c ExportAll2xlsx -d $OutputDir -f $Filename"


	echo "Running $RvtoolsExe with argument $Argumentlist"
	Start-Process -FilePath $RvtoolsExe -ArgumentList $ArgumentList -NoNewWindow -Wait
}

# Merge
$RvtoolsMergeExeArg = "-input $($FileList -join "";"" ) -output $OutputDir\$MergedOutputFile"
echo "Running $RvtoolsMergeExe with argment $RvtoolsMergeExeArg"
Start-Process -FilePath $RvtoolsMergeExe -ArgumentList $RvtoolsMergeExeArg -NoNewWindow -Wait

<# Send to SharePoint
$ctx = Connect-PnpOnline  -url $SharePointURL -CurrentCredentials

Get-ChildItem $OutputDir/RV*.xlsx |
	ForEach-Object { 
		echo "Upload $_"
		Add-PnpFile -Path $_ -Folder $SharePointPath
		echo "Move $_ to backup"
		Move-Item -Path $_ -Destination $OutputDir/backup
		echo "Done"
	}

Disconnect-PnPOnline -Connection $ctx	
#>

<# Idea: Send e-mail via Exchange EWS
https://github.com/bielawb/EWS

Connect-EWSService user@host
New-EWSMessage -To recipient -Subject subject -Body messge_body -Attachment attachment

#>

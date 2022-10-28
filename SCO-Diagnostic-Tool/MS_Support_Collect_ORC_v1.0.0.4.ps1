#NOTE: Special Credit to all team members for providing feedback and developing tips

#region auxiliar functions
function SelfElevate() {
    #got from http://www.expta.com/2017/03/how-to-self-elevate-powershell-script.html   and changed a bit
    if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
     if ([int](Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
      $CommandLine = "-File `"" + $Script:MyInvocation.MyCommand.Path + "`" " + $Script:MyInvocation.UnboundArguments
      Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
      Exit
     }
    }
}

# Recursive function to move all of the files that start with the File Name Prefix to the Directory To Move Files To.
function MoveFilesOutOfZipFileItems($shellItems, $directoryToMoveFilesToShell, $fileNamePrefix)
{
	# Loop through every item in the file/directory.
	foreach ($shellItem in $shellItems)
	{
		# If this is a directory, recursively call this function to iterate over all files/directories within it.
		if ($shellItem.IsFolder)
		{ 
			$totalItems += MoveFilesOutOfZipFileItems -shellItems $shellItem.GetFolder.Items() -directoryToMoveFilesTo $directoryToMoveFilesToShell -fileNameToMatch $fileNameToMatch
		}
		# Else this is a file.
		else
		{
			# If this file name starts with the File Name Prefix, move it to the specified directory.
			if ($shellItem.Name.StartsWith($fileNamePrefix))
			{
				$directoryToMoveFilesToShell.MoveHere($shellItem)
			}
		}			
	}
}

# Recursive function to move a directory into a Zip file, since we can move files out of a Zip file, but not directories, and copying a directory into a Zip file when it already exists is not allowed.
function MoveDirectoryIntoZipFile($parentInZipFileShell, $pathOfItemToCopy)
{
	# Get the name of the file/directory to copy, and the item itself.
	$nameOfItemToCopy = Split-Path -Path $pathOfItemToCopy -Leaf
	if ($parentInZipFileShell.IsFolder)
	{ $parentInZipFileShell = $parentInZipFileShell.GetFolder }
	$itemToCopyShell = $parentInZipFileShell.ParseName($nameOfItemToCopy)
	
	# If this item does not exist in the Zip file yet, or it is a file, move it over.
	if ($itemToCopyShell -eq $null -or !$itemToCopyShell.IsFolder)
	{
		$parentInZipFileShell.MoveHere($pathOfItemToCopy)
		
		# Wait for the file to be moved before continuing, to avoid erros about the zip file being locked or a file not being found.
		while (Test-Path -Path $pathOfItemToCopy)
		{ Start-Sleep -Milliseconds 10 }
	}
	# Else this is a directory that already exists in the Zip file, so we need to traverse it and copy each file/directory within it.
	else
	{
		# Copy each file/directory in the directory to the Zip file.
		foreach ($item in (Get-ChildItem -Path $pathOfItemToCopy -Force))
		{
			MoveDirectoryIntoZipFile -parentInZipFileShell $itemToCopyShell -pathOfItemToCopy $item.FullName
		}
	}
}

function Compress-ZipFile{
	[CmdletBinding()]
	param
	(
		[parameter(Position=1,Mandatory=$true)]
		[ValidateScript({Test-Path -Path $_})]
		[string]$FileOrDirectoryPathToAddToZipFile, 
	
		[parameter(Position=2,Mandatory=$false)]
		[string]$ZipFilePath,
		
		[Alias("Force")]
		[switch]$OverwriteWithoutPrompting
	)
	
	BEGIN { }
	END { }
	PROCESS
	{
		# If a Zip File Path was not given, create one in the same directory as the file/directory being added to the zip file, with the same name as the file/directory.
		if ($ZipFilePath -eq $null -or $ZipFilePath.Trim() -eq [string]::Empty)
		{ $ZipFilePath = Join-Path -Path $FileOrDirectoryPathToAddToZipFile -ChildPath '.zip' }
		
		# If the Zip file to create does not have an extension of .zip (which is required by the shell.application), add it.
		if (!$ZipFilePath.EndsWith('.zip', [StringComparison]::OrdinalIgnoreCase))
		{ $ZipFilePath += '.zip' }
		
		# If the Zip file to add the file to does not exist yet, create it.
		if (!(Test-Path -Path $ZipFilePath -PathType Leaf))
		{ New-Item -Path $ZipFilePath -ItemType File > $null }

		# Get the Name of the file or directory to add to the Zip file.
		$fileOrDirectoryNameToAddToZipFile = Split-Path -Path $FileOrDirectoryPathToAddToZipFile -Leaf

		# Get the number of files and directories to add to the Zip file.
		$numberOfFilesAndDirectoriesToAddToZipFile = (Get-ChildItem -Path $FileOrDirectoryPathToAddToZipFile -Recurse -Force).Count
		
		# Get if we are adding a file or directory to the Zip file.
		$itemToAddToZipIsAFile = Test-Path -Path $FileOrDirectoryPathToAddToZipFile -PathType Leaf

		# Get Shell object and the Zip File.
		$shell = New-Object -ComObject Shell.Application
		$zipShell = $shell.NameSpace($ZipFilePath)

		# We will want to check if we can do a simple copy operation into the Zip file or not. Assume that we can't to start with.
		# We can if the file/directory does not exist in the Zip file already, or it is a file and the user wants to be prompted on conflicts.
		$canPerformSimpleCopyIntoZipFile = $false

		# If the file/directory does not already exist in the Zip file, or it does exist, but it is a file and the user wants to be prompted on conflicts, then we can perform a simple copy into the Zip file.
		$fileOrDirectoryInZipFileShell = $zipShell.ParseName($fileOrDirectoryNameToAddToZipFile)
		$itemToAddToZipIsAFileAndUserWantsToBePromptedOnConflicts = ($itemToAddToZipIsAFile -and !$OverwriteWithoutPrompting)
		if ($fileOrDirectoryInZipFileShell -eq $null -or $itemToAddToZipIsAFileAndUserWantsToBePromptedOnConflicts)
		{
			$canPerformSimpleCopyIntoZipFile = $true
		}
		
		# If we can perform a simple copy operation to get the file/directory into the Zip file.
		if ($canPerformSimpleCopyIntoZipFile)
		{
			# Start copying the file/directory into the Zip file since there won't be any conflicts. This is an asynchronous operation.
			$zipShell.CopyHere($FileOrDirectoryPathToAddToZipFile)	# Copy Flags are ignored when copying files into a zip file, so can't use them like we did with the Expand-ZipFile function.
			
			# The Copy operation is asynchronous, so wait until it is complete before continuing.
			# Wait until we can see that the file/directory has been created.
			while ($zipShell.ParseName($fileOrDirectoryNameToAddToZipFile) -eq $null)
			{ Start-Sleep -Milliseconds 100 }
			
			# If we are copying a directory into the Zip file, we want to wait until all of the files/directories have been copied.
			if (!$itemToAddToZipIsAFile)
			{
				# Get the number of files and directories that should be copied into the Zip file.
				$numberOfItemsToCopyIntoZipFile = (Get-ChildItem -Path $FileOrDirectoryPathToAddToZipFile -Recurse -Force).Count
			
				# Get a handle to the new directory we created in the Zip file.
				$newDirectoryInZipFileShell = $zipShell.ParseName($fileOrDirectoryNameToAddToZipFile)
				
				# Wait until the new directory in the Zip file has the expected number of files and directories in it.
				while ((GetNumberOfItemsInZipFileItems -shellItems $newDirectoryInZipFileShell.GetFolder.Items()) -lt $numberOfItemsToCopyIntoZipFile)
				{ Start-Sleep -Milliseconds 100 }
			}
		}
		# Else we cannot do a simple copy operation. We instead need to move the files out of the Zip file so that we can merge the directory, or overwrite the file without the user being prompted.
		# We cannot move a directory into the Zip file if a directory with the same name already exists, as a MessageBox warning is thrown, not a conflict resolution prompt like with files.
		# We cannot silently overwrite an existing file in the Zip file, as the flags passed to the CopyHere/MoveHere functions seem to be ignored when copying into a Zip file.
		else
		{
			# Create a temp directory to hold our file/directory.
			$tempDirectoryPath = $null
			$tempDirectoryPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
			New-Item -Path $tempDirectoryPath -ItemType Container > $null
		
			# If we will be moving a directory into the temp directory.
			$numberOfItemsInZipFilesDirectory = 0
			if ($fileOrDirectoryInZipFileShell.IsFolder)
			{
				# Get the number of files and directories in the Zip file's directory.
				$numberOfItemsInZipFilesDirectory = GetNumberOfItemsInZipFileItems -shellItems $fileOrDirectoryInZipFileShell.GetFolder.Items()
			}
		
			# Start moving the file/directory out of the Zip file and into a temp directory. This is an asynchronous operation.
			$tempDirectoryShell = $shell.NameSpace($tempDirectoryPath)
			$tempDirectoryShell.MoveHere($fileOrDirectoryInZipFileShell)
			
			# If we are moving a directory, we need to wait until all of the files and directories in that Zip file's directory have been moved.
			$fileOrDirectoryPathInTempDirectory = Join-Path -Path $tempDirectoryPath -ChildPath $fileOrDirectoryNameToAddToZipFile
			if ($fileOrDirectoryInZipFileShell.IsFolder)
			{
				# The Move operation is asynchronous, so wait until it is complete before continuing. That is, sleep until the Destination Directory has the same number of files as the directory in the Zip file.
				while ((Get-ChildItem -Path $fileOrDirectoryPathInTempDirectory -Recurse -Force).Count -lt $numberOfItemsInZipFilesDirectory)
				{ Start-Sleep -Milliseconds 100 }
			}
			# Else we are just moving a file, so we just need to check for when that one file has been moved.
			else
			{
				# The Move operation is asynchronous, so wait until it is complete before continuing.
				while (!(Test-Path -Path $fileOrDirectoryPathInTempDirectory))
				{ Start-Sleep -Milliseconds 100 }
			}
			
			# We want to copy the file/directory to add to the Zip file to the same location in the temp directory, so that files/directories are merged.
			# If we should automatically overwrite files, do it.
			if ($OverwriteWithoutPrompting)
			{ Copy-Item -Path $FileOrDirectoryPathToAddToZipFile -Destination $tempDirectoryPath -Recurse -Force }
			# Else the user should be prompted on each conflict.
			else
			{ Copy-Item -Path $FileOrDirectoryPathToAddToZipFile -Destination $tempDirectoryPath -Recurse -Confirm -ErrorAction SilentlyContinue }	# SilentlyContinue errors to avoid an error for every directory copied.

			# For whatever reason the zip.MoveHere() function is not able to move empty directories into the Zip file, so we have to put dummy files into these directories 
			# and then remove the dummy files from the Zip file after.
			# If we are copying a directory into the Zip file.
			$dummyFileNamePrefix = 'Dummy.File'
			[int]$numberOfDummyFilesCreated = 0
			if ($fileOrDirectoryInZipFileShell.IsFolder)
			{
				# Place a dummy file in each of the empty directories so that it gets copied into the Zip file without an error.
				$emptyDirectories = Get-ChildItem -Path $fileOrDirectoryPathInTempDirectory -Recurse -Force -Directory | Where-Object { (Get-ChildItem -Path $_ -Force) -eq $null }
				foreach ($emptyDirectory in $emptyDirectories)
				{
					$numberOfDummyFilesCreated++
					New-Item -Path (Join-Path -Path $emptyDirectory.FullName -ChildPath "$dummyFileNamePrefix$numberOfDummyFilesCreated") -ItemType File -Force > $null
				}
			}		

			# If we need to copy a directory back into the Zip file.
			if ($fileOrDirectoryInZipFileShell.IsFolder)
			{
				MoveDirectoryIntoZipFile -parentInZipFileShell $zipShell -pathOfItemToCopy $fileOrDirectoryPathInTempDirectory
			}
			# Else we need to copy a file back into the Zip file.
			else
			{
				# Start moving the merged file back into the Zip file. This is an asynchronous operation.
				$zipShell.MoveHere($fileOrDirectoryPathInTempDirectory)
			}
			
			# The Move operation is asynchronous, so wait until it is complete before continuing.
			# Sleep until all of the files have been moved into the zip file. The MoveHere() function leaves empty directories behind, so we only need to watch for files.
			do
			{
				Start-Sleep -Milliseconds 100
				$files = Get-ChildItem -Path $fileOrDirectoryPathInTempDirectory -Force -Recurse | Where-Object { !$_.PSIsContainer }
			} while ($files -ne $null)
			
			# If there are dummy files that need to be moved out of the Zip file.
			if ($numberOfDummyFilesCreated -gt 0)
			{
				# Move all of the dummy files out of the supposed-to-be empty directories in the Zip file.
				MoveFilesOutOfZipFileItems -shellItems $zipShell.items() -directoryToMoveFilesToShell $tempDirectoryShell -fileNamePrefix $dummyFileNamePrefix
				
				# The Move operation is asynchronous, so wait until it is complete before continuing.
				# Sleep until all of the dummy files have been moved out of the zip file.
				do
				{
					Start-Sleep -Milliseconds 100
					[Object[]]$files = Get-ChildItem -Path $tempDirectoryPath -Force -Recurse | Where-Object { !$_.PSIsContainer -and $_.Name.StartsWith($dummyFileNamePrefix) }
				} while ($files -eq $null -or $files.Count -lt $numberOfDummyFilesCreated)
			}
			
			# Delete the temp directory that we created.
			Remove-Item -Path $tempDirectoryPath -Force -Recurse > $null
		}
	}
}
function MakeNewZipFile($source,$archive) { #https://stackoverflow.com/questions/40692024/zip-and-unzip-file-in-powershell-4
    Add-Type -assembly "system.io.compression.filesystem"
    [io.compression.zipfile]::CreateFromDirectory($source, $archive)
}
function AppendOutputToFileInTargetFolder($obj, $fileName) {
    $resultFilePath = Join-Path -Path $resultFolder -ChildPath $fileName    
    if (!(Test-Path $resultFilePath))
    {
       New-Item $resultFilePath -ItemType File -Force | Out-Null
    }
    $obj | Out-File -FilePath $resultFilePath -Encoding utf8 -Append 
}
function CopyFileToTargetFolder($fileName, $subFolderName) {
  if ([string]::IsNullOrEmpty($subFolderName) -or $subFolderName -eq ".") { 
    Copy-Item $fileName -Destination $resultFolder}
  else  {
    New-Item -ItemType Directory -Force -Path "$resultFolder\$subFolderName" | Out-Null
    Copy-Item $fileName -Destination "$resultFolder\$subFolderName" }
}
function CreateNewFileInTargetFolder($fileName) {
    New-Item -Force -ItemType File -Path $resultFolder -Name $fileName | Out-Null
}
function DeleteFileInTargetFolder($fileName) {
    return Remove-Item -Path (Join-Path -Path $resultFolder -ChildPath $fileName)
}
function CreateNewFolderInTargetFolder($folderName) {
    New-Item -Force -ItemType Directory -Path $resultFolder -Name $folderName | Out-Null
}
function GetFileNameInTargetFolder($fileName) {
    return Join-Path -Path $resultFolder -ChildPath $fileName
}
function GetFileContentInTargetFolder($fileName) {
    return Get-Content -Path (Join-Path -Path $resultFolder -ChildPath $fileName) | Out-String
    #return [IO.File]::ReadAllText( (Join-Path -Path $resultFolder -ChildPath $fileName) )
}
function AppendOutputToCsvFileInTargetFolder($dataTable, $fileName) {     
    $resultFilePath = Join-Path -Path $resultFolder -ChildPath $fileName
    if ($dataTable.Rows.Count -eq 0) 
    {
        $header = ""
        foreach ($col in $dataTable.Columns) {
            $header += $col.ColumnName +","
        }
        $header = $header.Remove($header.Length-1,1) 
        AppendOutputToFileInTargetFolder $header $fileName
    } 
    else 
    {
        #$dataTable | export-csv -Path $resultFilePath -Encoding UTF8 -Append -NoTypeInformation
        AppendOutputToFileInTargetFolder ($dataTable | ConvertTo-Csv -NoTypeInformation) $fileName
    }
}
function Try-Invoke-SqlCmd{
    param (
            [Parameter(Mandatory=$true)] [string]$SQLInstance,
            [Parameter(Mandatory=$true)] [string]$SQLDatabase,
            [Parameter(Mandatory=$true)] [string]$Query
    )
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server=$SQLInstance; Database=$SQLDatabase; Trusted_Connection=True"
    $SqlConnection.Open() 

    $SqlAdp = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = $SQLQuery
    $SqlCmd.Connection = $SqlConnection
    $SqlCmd.CommandTimeout = 0
    $SqlAdp.SelectCommand = $SqlCmd

#    if ($CaptureSqlPrintMessages) {
        $global:SqlPrintMessages=""
        $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {
            param($sender, $event) 
            $global:SqlPrintMessages += "`n" + $event.Message 
        };
        $SqlConnection.add_InfoMessage($handler); 
        $SqlConnection.FireInfoMessageEventOnUserErrors = $true;
#    }

    $DS = New-Object System.Data.DataSet
    $SqlAdp.Fill($DS) | out-null  # keep the out-null otherwise $DS will return as Object[]
    return $DS;
}
function SaveSQLResultToFile($dataTable, $fileName, $batch) {  #, $includeBatchInResultSet) {
    $TempFileName = ([guid]::NewGuid()).ToString()    
    AppendOutputToCsvFileInTargetFolder ($dataTable) $TempFileName
    #if ($includeBatchInResultSet -eq $null) {$includeBatchInResultSet=$true}
    #if ($includeBatchInResultSet) {
    #    AppendOutputToFileInTargetFolder "" $TempFileName
    #    AppendOutputToFileInTargetFolder "/*------------------------`r`n$($batch.Trim())`r`n------------------------*/" $TempFileName 
    #}
    AppendOutputToFileInTargetFolder "" $fileName
    AppendOutputToFileInTargetFolder (GetFileContentInTargetFolder $TempFileName) $fileName
    DeleteFileInTargetFolder $TempFileName
}
function SaveSQLResultSetsToFiles($SQLInstance, $SQLDatabase, $SQLQuery, $fileName) { #, $includeBatchInResultSet) {
    if ([string]::IsNullOrEmpty($script:SQLResultSetCounter)) {
		$script:SQLResultSetCounter=1
	}
    $batches = $SQLQuery -split '(?:\bGO\b)'
    foreach($batch in $batches) {
        if ([string]::IsNullOrEmpty($batch.Trim())) {continue}
        $DS = Try-Invoke-SqlCmd -SQLInstance $SQLInstance -SQLDatabase $SQLDatabase -Query $batch 
        $targetCsvFileName = $fileName 

#        if ($DS.Tables.Count -eq 0) { # write the Sql Print messages
#            AppendOutputToFileInTargetFolder $global:SqlPrintMessages $targetCsvFileName
#        }
#        else {
            foreach($dataTable in $DS.Tables) 
            {            
                if ([string]::IsNullOrEmpty($targetCsvFileName))
                {
                    $targetCsvFileName = "SQLResultSet_$script:SQLResultSetCounter.csv"
                    $script:SQLResultSetCounter++
                }            
                SaveSQLResultToFile ($dataTable) $targetCsvFileName $batch #$includeBatchInResultSet      
            }
            AppendOutputToFileInTargetFolder $global:SqlPrintMessages $targetCsvFileName

#       }

        if ($includeBatchInResultSet -eq $null) {
			$includeBatchInResultSet=$true
		}
        if ($includeBatchInResultSet) {
            AppendOutputToFileInTargetFolder "" $targetCsvFileName
            AppendOutputToFileInTargetFolder "/*------------------------`r`n$($batch.Trim())`r`n------------------------*/" $targetCsvFileName 
        }
    }
}
#endregion


function exportRegKeys {
    createNewFolderInTargetFolder "TLS"
    Reg export "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders" (GetFileNameInTargetFolder "TLS\SecurityProviders.txt")| Out-Null
    Reg export "HKLM\SOFTWARE\Microsoft\.NETFramework" (GetFileNameInTargetFolder "TLS\NETFramework.txt")| Out-Null
    Reg export "HKLM\SOFTWARE\WOW6432Node\Microsoft\.NETFramework" (GetFileNameInTargetFolder "TLS\WOW6432Node_NETFramework.txt")| Out-Null

    Reg export "HKLM\SOFTWARE\Microsoft\Silverlight" (GetFileNameInTargetFolder "Silverlight.txt")| Out-Null
    Reg export "HKLM\SOFTWARE\Microsoft\Microsoft OLE DB Driver for SQL Server" (GetFileNameInTargetFolder "TLS\Microsoft OLE DB Driver for SQL Server.txt")| Out-Null

    Reg export "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" (GetFileNameInTargetFolder "TLS\DotNetFwV3.5.txt")| Out-Null
    Reg export "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4"   (GetFileNameInTargetFolder "TLS\DotNetFwV4.txt")| Out-Null
}

function exportEventLog {
    (get-wmiobject win32_nteventlogfile -filter "logfilename = 'System'").BackupEventlog((GetFileNameInTargetFolder "System.evtx")) | Out-Null
    wevtutil archive-log (GetFileNameInTargetFolder "System.evtx") /l:en-US | Out-Null

    (get-wmiobject win32_nteventlogfile -filter "logfilename = 'Application'").BackupEventlog((GetFileNameInTargetFolder "Application.evtx")) | Out-Null
    wevtutil archive-log (GetFileNameInTargetFolder "Application.evtx") /l:en-US | Out-Null

    wevtutil epl Microsoft-Windows-WMI-Activity/Operational (GetFileNameInTargetFolder "Microsoft-Windows-WMI-Activity_Operational.evtx") | Out-Null

}

function getServices {
    #(Get-Service |  Sort-Object DisplayName| select-object Status, DisplayName, Name |ft ) >> (GetFileNameInTargetFolder "Services.txt")
	(gwmi win32_service | Select-Object Name, StartMode, State, DisplayName, StartName, pathname) >> (GetFileNameInTargetFolder "Services.txt")
}

function getProcesses {
   AppendOutputToFileInTargetFolder ( Get-ExecutionPolicy -List ) Get-ExecutionPolicy.txt

    if ( $PSVersionTable.PSVersion.Major -ge 3 ) { 
        $processWithAllInfo = Get-Process -IncludeUserName | ? {$_.Id -ne 0 }| select *,CurrentCPU 
    }
    else {
        $processWithAllInfo = Get-Process | ? {$_.Id -ne 0 }| select *,CurrentCPU 
    }

    $PID_CurrentCPU=Get-WmiObject Win32_PerfFormattedData_PerfProc_Process | ? {$_.IDProcess -ne 0 } | select IDProcess, @{ Name = 'PercentProcessorTime';  Expression = {$_.PercentProcessorTime / ($env:NUMBER_OF_PROCESSORS) }} 
    foreach($p in $processWithAllInfo) {  
        $p.CurrentCPU = ( $PID_CurrentCPU | ? {$_.IDProcess -eq $p.Id}  ).PercentProcessorTime 
    }  
    AppendOutputToFileInTargetFolder (  $processWithAllInfo | select Handles, WS, CurrentCPU, Id, UserName, ProcessName | ft -Wrap ) Get-Process_WithCurrentCPU.txt
    AppendOutputToFileInTargetFolder (  $processWithAllInfo ) Get-Process_WithAllDetails.txt
    AppendOutputToFileInTargetFolder (  $processWithAllInfo | ? {$_.Name -eq "System" } | Select StartTime ) MachineStartTime.txt
    AppendOutputToFileInTargetFolder (  $processWithAllInfo | ? {$_.Id -ne 0  -and  $_.CurrentCPU -gt 0} | select Id,Name,CurrentCPU )  Get-Process_OnlyActiveOnes.txt
    AppendOutputToFileInTargetFolder (  $processWithAllInfo | ? {$_.Id -ne 0} | Sort-Object -Property CurrentCPU -Descending | select Id,Name,CurrentCPU -First 10  )  Get-Process_Top10_ByCPU.txt
    AppendOutputToFileInTargetFolder (  $processWithAllInfo | ? {$_.Id -ne 0} | Sort-Object -Property WS -Descending | select Id,Name,WS -First 10  )  Get-Process_Top10_ByWorkingSet.txt
}

function exportMSInfo32 {
    #$msinfo32= Start-Process msinfo32.exe -PassThru -ArgumentList "/nfo ""$((GetFileNameInTargetFolder "$env:computername.nfo"))"""
    $msinfo32= Start-Process msinfo32.exe -PassThru -ArgumentList "/report ""$((GetFileNameInTargetFolder "msinfo32.txt"))"""
    AppendOutputToFileInTargetFolder (dir env:* | ConvertTo-Csv -NoTypeInformation) "EnvVars.csv"
    AppendOutputToFileInTargetFolder (Get-HotFix | Format-table -Wrap -AutoSize)  "Get-Hotfix.txt"
    AppendOutputToFileInTargetFolder ($PSVersionTable) "PSVersionTable.txt"
    AppendOutputToFileInTargetFolder ($PSVersionTable.PSCompatibleVersions) "PSCompatibleVersions.txt"   
    AppendOutputToFileInTargetFolder (whoami) "Whoami.txt"
    AppendOutputToFileInTargetFolder (netstat /abof) "netstat_abof.txt"
    AppendOutputToFileInTargetFolder (Get-WmiObject -Class Win32_Product | Select-Object Version, Name, InstallDate ) "ProgramVersions.txt"
    AppendOutputToFileInTargetFolder (Invoke-Expression -Command "dism.exe /online /Get-intl") "LanguageInfo.txt"

    AppendOutputToFileInTargetFolder ( (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1mb ) TotalRAM.txt

	return $msinfo32
}


#region Queries from OrchestratorDB

function runSQLQueries ($SQLInstance, $SQLDatabase) {
    if ([string]::IsNullOrEmpty($SQLDatabase)){
        $SQLDatabase = 'Orchestrator'
    }
	$SQL_Queries =@{}
	
    $SQL_Queries['SQL_dbo.VERSION']=@'
SELECT * FROM [dbo].[VERSION] 
'@

    $SQL_Queries['SQL_dbo.EVENTS']=@'
SELECT * FROM [dbo].[EVENTS] order by DateTime desc 
'@

   $SQL_Queries['SQL_dbo.ACTIONSERVERS']=@'
SELECT * FROM [dbo].[ACTIONSERVERS]
'@

   $SQL_Queries['SQL_MaintenanceTasks']=@'
SELECT * FROM [Microsoft.SystemCenter.Orchestrator.Maintenance].[MaintenanceTasks] 
'@

   $SQL_Queries['SQL_LogPurgeSettings']=@'
SELECT DataName, DataValue FROM [dbo].[CONFIGURATION] where TypeGUID = 'F05CF395-E7D0-4805-A8DC-588FE9C3E4C9'
'@

   $SQL_Queries['SQL_dbo.POLICY_PUBLISH_QUEUE']=@'
Select Count(*) From POLICYINSTANCES WITH (NOLOCK)      
Select Count(*) From OBJECTINSTANCES WITH (NOLOCK)     
Select Count(*) From OBJECTINSTANCEDATA WITH (NOLOCK)   
Select Count(*) From EVENTS WITH (NOLOCK)               
Select Count(*) From POLICY_PUBLISH_QUEUE WITH (NOLOCK) 
'@

   $SQL_Queries['SQL_OrphanRunbooks']=@'
Select Count (*)
from [dbo].[POLICYINSTANCES] AS pinst, [dbo].[POLICY_REQUEST_HISTORY] AS prq
where pinst.[PolicyID] = prq.[PolicyID] AND pinst.[SeqNumber] = prq.[SeqNumber] AND pinst.[TimeEnded] IS NULL AND prq.[Active] = 0 
select pinst.[UniqueID],pinst.[PolicyID]
from [dbo].[POLICYINSTANCES] AS pinst, [dbo].[POLICY_REQUEST_HISTORY] AS prq
where pinst.[PolicyID] = prq.[PolicyID] AND pinst.[SeqNumber] = prq.[SeqNumber] AND pinst.[TimeEnded] IS NULL AND prq.[Active] = 0 
'@

  $SQL_Queries['SQL_FailedRunbooksAndActivities']=@'
SELECT P.Name as [Runbook Name], ACT.name as [Activity Name], OBI.ObjectStatus as [Activity Status], OBID.[Value] as [Error Summary], OBI.StartTime, OBI.EndTime, *
FROM [Microsoft.SystemCenter.Orchestrator].[Activities] ACT
-- to get Runbook properties of the activity
inner join POLICIES P on ACT.RunbookId = P.UniqueID
-- to get activity execution status history 
inner join OBJECTINSTANCES OBI on ACT.ID = OBI.ObjectID 
--to get key and value of of each activity history e.g. error summary
inner join OBJECTINSTANCEDATA OBID on OBID.ObjectInstanceID = OBI.UniqueID 
where  OBID.[Key] = 'ErrorSummary.Text' and  OBI.ObjectStatus = 'Failed' 
order by  P.Name, ACT.Name, OBI.StartTime desc
'@


 $SQL_Queries['SQL_Date']=@'
select SYSDATETIMEOFFSET() as LocalTime,GETUTCDATE() as UtcTime
'@

 $SQL_Queries['SQL_EncryptionKeys']=@'
select top 10 @@SERVERNAME, * from sys.symmetric_keys;
select top 10 @@SERVERNAME, * from sys.asymmetric_keys;
select name, database_id, is_master_key_encrypted_by_server from sys.databases Where is_master_key_encrypted_by_server=1;
use master; select top 10 @@SERVERNAME, * from sys.symmetric_keys;
use master; select top 10 @@SERVERNAME, * from sys.asymmetric_keys;
'@
#--use master go select top 10 @@SERVERNAME, * from sys.symmetric_keys
#--use master go select top 10 @@SERVERNAME, * from sys.asymmetric_keys


    $SQL_Queries['SQL_Databases']=@'
SELECT name,is_broker_enabled,compatibility_level,recovery_model_desc,* FROM sys.databases order by 1
'@

    $SQL_Queries['SQL_dm_os_schedulers']=@'
SELECT * FROM sys.dm_os_schedulers WHERE scheduler_id < 255;
'@

    $SQL_Queries['SQL_CurrentlyRunningQueries']=@'
SELECT SUBSTRING(sqltext.text, ( req.statement_start_offset / 2 ) + 1, 
              ( ( CASE WHEN req.statement_end_offset <= 0
                       THEN DATALENGTH(sqltext.text) 
              ELSE req.statement_end_offset END - 
       req.statement_start_offset ) / 2 ) + 1) AS statement_text,
sqltext.TEXT, req.last_wait_type,req.session_id,req.status,req.command,req.cpu_time,req.total_elapsed_time,blocking_session_id
,database_id,DB_NAME(database_id), p.hostname,p.hostprocess
FROM sys.dm_exec_requests req
CROSS APPLY sys.dm_exec_sql_text(sql_handle) AS sqltext 
left join sys.sysprocesses p on req.session_id=p.spid
where req.session_id !=@@spid and  req.last_wait_type not like '%broker%'
'@

    $SQL_Queries['SQL_database_scoped_configurations_IfGe2016']=@'
if (select convert(smallint,SERVERPROPERTY('ProductMajorVersion'))) >= 13 --greater than or equal sql 2016
SELECT
(select value from sys.database_scoped_configurations as dsc where dsc.name = 'MAXDOP') AS [MaxDop],
(select value_for_secondary from sys.database_scoped_configurations as dsc where dsc.name = 'MAXDOP') AS [MaxDopForSecondary],
(select value from sys.database_scoped_configurations as dsc where dsc.name = 'LEGACY_CARDINALITY_ESTIMATION') AS [LegacyCardinalityEstimation],
(select ISNULL(value_for_secondary, 2) from sys.database_scoped_configurations as dsc where dsc.name = 'LEGACY_CARDINALITY_ESTIMATION') AS [LegacyCardinalityEstimationForSecondary],
(select value from sys.database_scoped_configurations as dsc where dsc.name = 'PARAMETER_SNIFFING') AS [ParameterSniffing],
(select ISNULL(value_for_secondary, 2) from sys.database_scoped_configurations as dsc where dsc.name = 'PARAMETER_SNIFFING') AS [ParameterSniffingForSecondary],
(select value from sys.database_scoped_configurations as dsc where dsc.name = 'QUERY_OPTIMIZER_HOTFIXES') AS [QueryOptimizerHotfixes],
(select ISNULL(value_for_secondary, 2) from sys.database_scoped_configurations as dsc where dsc.name = 'QUERY_OPTIMIZER_HOTFIXES') AS [QueryOptimizerHotfixesForSecondary]
else
select 'no sys.database_scoped_configurations available for this sql version'
'@

    $SQL_Queries['SQL_DatabaseFiles']=@'
select sys.databases.name, sys.databases.database_id,sys.master_files.physical_name,size*8/1024 SizeInMB  from sys.databases join sys.master_files on sys.databases.database_id = sys.master_files.database_id where sys.databases.source_database_id is null
'@

    $SQL_Queries['SQL_sp_configure']=@'
exec sp_configure 'show advanced options',1 
RECONFIGURE
exec sp_configure
'@

    $SQL_Queries['SQL_dm_os_sys_info']=@'
select * from sys.dm_os_sys_info 
'@

    $SQL_Queries['SQL_dm_os_wait_stats']=@'
SELECT TOP 15 * FROM sys.dm_os_wait_stats ORDER BY wait_time_ms DESC
'@

    $SQL_Queries['SQL_LoginsInfo']=@'
select name,language,sysadmin from sys.syslogins order by 1
'@

    $SQL_Queries['SQL_DbUsersInfo']=@'
DECLARE @DB_USers TABLE
(DBName sysname, UserName sysname, LoginType sysname, AssociatedRole varchar(max),create_date datetime,modify_date datetime)
INSERT @DB_USers
EXEC sp_MSforeachdb
'use [?]
SELECT ''?'' AS DB_Name,
case prin.name when ''dbo'' then prin.name + '' (''+ (select SUSER_SNAME(owner_sid) from master.sys.databases where name =''?'') + '')'' else prin.name end AS UserName,
prin.type_desc AS LoginType,
isnull(USER_NAME(mem.role_principal_id),'''') AS AssociatedRole ,create_date,modify_date
FROM sys.database_principals prin
LEFT OUTER JOIN sys.database_role_members mem ON prin.principal_id=mem.member_principal_id
WHERE prin.sid IS NOT NULL and prin.sid NOT IN (0x00) and
prin.is_fixed_role <> 1 AND prin.name NOT LIKE ''##%'''
SELECT
dbname,username ,logintype ,create_date ,modify_date ,
STUFF(
	(SELECT ',' + CONVERT(VARCHAR(500),associatedrole)
	FROM @DB_USers user2
	WHERE
	user1.DBName=user2.DBName AND user1.UserName=user2.UserName
	FOR XML PATH('')
	)
	,1,1,''
	) AS Permissions_user
FROM @DB_USers user1
WHERE dbname=DB_NAME()
GROUP BY dbname,username ,logintype ,create_date ,modify_date
ORDER BY DBName,username
'@

    $SQL_Queries['SQL_FragmentationInfo']=@'
SELECT OBJECT_NAME(ind.OBJECT_ID) AS TableName,
ind.name AS IndexName, indexstats.index_type_desc AS IndexType,
indexstats.avg_fragmentation_in_percent--,*
FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, NULL) indexstats
INNER JOIN sys.indexes ind
ON ind.object_id = indexstats.object_id
AND ind.index_id = indexstats.index_id
ORDER BY indexstats.avg_fragmentation_in_percent DESC
'@

    $SQL_Queries['SQL_TableSizeInfo']=@'
declare c cursor local FORWARD_ONLY READ_ONLY for
select '['+ s.name +'].['+ o.name +']'
from sys.objects o
inner join sys.schemas s on o.schema_id=s.schema_id 
where o.type='U' 
order by o.name
declare @fqName nvarchar(max)
declare  @tbl table(
name nvarchar(max),
rows bigint,
reserved varchar(18),
data varchar(18),
index_size varchar(18),
unused varchar(18)
)
open c
while 1=1
begin
fetch c into @fqName
if @@FETCH_STATUS<>0 break
	insert into @tbl
	exec sp_spaceused @fqName
end
close c
deallocate c
select name,rows,data,index_size,unused from @tbl order by rows desc
'@

    $SQL_Queries['SQL_Info'] = @'
select @@VERSION as "@@VERSION"
create table #SVer(ID int,  Name  sysname, Internal_Value int, Value nvarchar(512))
insert #SVer exec master.dbo.xp_msver
if exists (select 1 from sys.all_objects where name = 'dm_os_host_info' and type = 'V' and is_ms_shipped = 1)
begin
insert #SVer select t.*
from sys.dm_os_host_info
CROSS APPLY (
VALUES
(1001, 'host_platform', 0, host_platform),
(1002, 'host_distribution', 0, host_distribution),
(1003, 'host_release', 0, host_release),
(1004, 'host_service_pack_level', 0, host_service_pack_level),
(1005, 'host_sku', host_sku, '')
) t(id, [name], internal_value, [value])
end
declare @SmoRoot nvarchar(512)
exec master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'SOFTWARE\Microsoft\MSSQLServer\Setup', N'SQLPath', @SmoRoot OUTPUT
SELECT
(select Value from #SVer where Name = N'ProductName') AS [Product],
SERVERPROPERTY(N'ProductVersion') AS [VersionString],
(select Value from #SVer where Name = N'Language') AS [Language],
(select Value from #SVer where Name = N'Platform') AS [Platform],
CAST(SERVERPROPERTY(N'Edition') AS sysname) AS [Edition],
(select Internal_Value from #SVer where Name = N'ProcessorCount') AS [Processors],
(select Value from #SVer where Name = N'WindowsVersion') AS [OSVersion],
(select Internal_Value from #SVer where Name = N'PhysicalMemory') AS [PhysicalMemory],
CAST(ISNULL(SERVERPROPERTY('IsClustered'),N'') AS bit) AS [IsClustered],
@SmoRoot AS [RootDirectory],
convert(sysname, serverproperty(N'collation')) AS [Collation],
( select Value from #SVer where Name =N'host_platform') AS [HostPlatform],
( select Value from #SVer where Name =N'host_release') AS [HostRelease],
( select Value from #SVer where Name =N'host_service_pack_level') AS [HostServicePackLevel],
( select Value from #SVer where Name =N'host_distribution') AS [HostDistribution]
drop table #SVer
'@

	$SQL_Queries['SQL_Dbcc_Useroptions'] = @'
dbcc useroptions
'@

	$SQL_Queries['SQL_information_schema_columns'] = @'
select * from information_schema.columns order by Table_name,COLUMN_NAME
'@

	$SQL_Queries['SQL_BackupInfo'] = @'
SELECT 
    database_name
    , case type
	when 'D' then 'Database'
	when 'I' then 'Differential database'
	when 'L' then 'Log'
	when 'F' then 'File or filegroup'
	when 'G' then 'Differential file'
	when 'P' then 'Partial'
	when 'Q' then 'Differential partial'
	else '(unknown)'
	 end AS BackupType
    , MAX(backup_start_date) AS LastBackupDate
    , GETDATE() AS CurrentDate
    , DATEDIFF(DD,MAX(backup_start_date),GETDATE()) AS DaysSinceBackup
FROM msdb.dbo.backupset BS JOIN master.dbo.sysdatabases SD ON BS.database_name = SD.[name]
GROUP BY database_name, type 
ORDER BY database_name, type
'@

    foreach($SQL_Query in $SQL_Queries.Keys) {        
	    SaveSQLResultSetsToFiles $SQLInstance $SQLDatabase ($SQL_Queries[$SQL_Query]) "$SQL_Query.csv"    
	} 	
}
#endregion


function exportLogFiles{
    Get-ChildItem -Path "$env:SystemDrive\ProgramData\Microsoft System Center 2012\Orchestrator" -Include *.log -Recurse -Force -ErrorAction SilentlyContinue | % {
        $UpDir1 = $_.Directory;
        $UpDir2 = $UpDir1.Parent;
        $UpDir3 = $UpDir2.Parent;
        if ( $UpDir1.Name -eq "Logs") {
            $userFolderName=$UpDir2;
            CopyFileToTargetFolder $_.FullName "ORCH_LogFiles\$userFolderName"
        }    
    } -ErrorAction SilentlyContinue

    Get-ChildItem -Path "$env:SystemDrive\Users\*\AppData\Local\Microsoft System Center 2012\Orchestrator\LOGS" -Include *.log -Recurse -Force -ErrorAction SilentlyContinue | % {
       CopyFileToTargetFolder $_.FullName "ORCH_InstallationLogFiles"
    } -ErrorAction SilentlyContinue
}

function exportIISsettings{
	$appPools = Get-IISAppPool
	AppendOutputToFileInTargetFolder ($appPools) "IISsettings.txt"
	foreach($appPool in $appPools){
	   AppendOutputToFileInTargetFolder ($appPool.Name) "IISsettings.txt"
	   AppendOutputToFileInTargetFolder ($appPool.ProcessModel.UserName) "IISsettings.txt"
	}

	$sites = Get-IISSite 
	AppendOutputToFileInTargetFolder ($sites) "IISsettings.txt"
	foreach ($site in $sites){
		AppendOutputToFileInTargetFolder ($site.Name) "IISsettings.txt"
		AppendOutputToFileInTargetFolder (Get-IISSiteBinding $site.Name) "IISsettings.txt"
	}
	
}


function loadOrchDB{
	try {
		[System.Reflection.Assembly]::LoadWithPartialName("System.Security") | out-null
		$settingsFile = Join-Path "${Env:CommonProgramFiles(x86)}" "Microsoft System Center 2012\Orchestrator\Shared\Settings.dat"
		if ((Test-Path $settingsFile) -eq $false){
            $settingsFile = Join-Path "${Env:CommonProgramFiles}" "Microsoft System Center 2012\Orchestrator\Shared\Settings.dat"
            if ((Test-Path $settingsFile) -eq $false){
                Write-Error "No permission or Settings.dat file not found at: $settingsFile"
			    break
            }
		}
		$encryptedData = get-content -encoding byte $settingsFile
		$decryptedData = [System.Security.Cryptography.ProtectedData]::Unprotect($encryptedData, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
		$decryptedText = [System.Text.Encoding]::Unicode.GetString($decryptedData ) -replace "`r`n.`$",""
		#<Configuration><Server>sql</Server><Provider>MSOLEDBSQL</Provider><Login/><Database>Orchestrator</Database></Configuration>

		$xmlDoc = [System.Xml.XmlDocument]::new()
		$xmlDoc.LoadXml($decryptedText)
		$SQLInstance = $xmlDoc.Configuration.Server
		$SQLDatabase = $xmlDoc.Configuration.Database
		
        $myArray = @()
        $myArray += $SQLInstance
        $myArray += $SQLDatabase
        return $myArray
	}
	catch
	{
		"Error decrypting settings.dat file: Error Exception is + $_.Exception " 
		   Throw $_.Exception 
	} 
}


############################# MAIN function #####################################
	$resultPrefix = "ORCH"
    $isElevationRequired = $true # hardcoded for now
    if ($isElevationRequired) {SelfElevate}
    $resultFolderPath = $env:TEMP #Get-Location 
    $resultDateTime = (Get-Date).ToString("yyyy-MM-dd__HH.mm.fff")    
	$resultFolder = New-Item -Force -ItemType Directory -Path $resultFolderPath -Name "$($resultPrefix)_$resultDateTime"
	Start-Transcript -Path "$resultFolder\Transcript_$resultDateTime.txt" -NoClobber | Out-Null
    Write-Host "This script does *NOT* make any change in your ORCH environment. It is completely read-only."
    Write-Host ""
    Write-Host "Collection started at $resultDateTime. (local time)"
    Write-Host "Please wait for completion. This can take a few minutes..." -ForegroundColor Yellow
    Write-Host "(Please ignore any Warning and Errors)"

    Write-Host "_______________________________________"
  
    Write-Host "Gathering Registry Keys on local server..." -ForegroundColor Cyan
    exportRegKeys
    Write-Host "Gathering EventLog on local server..." -ForegroundColor Cyan
    exportEventLog
    Write-Host "Gathering Services Information on local server..." -ForegroundColor Cyan
    getServices
    Write-Host "Gathering Processes Information on local server..." -ForegroundColor Cyan
    getProcesses
    Write-Host "Exporting MSInfo32 on local server..." -ForegroundColor Cyan
    $msinfo32 = exportMSInfo32
	Write-Host "Exporting IIS settings on local server..." -ForegroundColor Cyan
	exportIISsettings
    Write-Host "Getting automatic SQL Server connection info" -ForegroundColor Cyan	
    $array = loadOrchDB
    $SQLInstance = $array[0]
    $SQLDatabase = $array[1]
    if ($SQLInstance -eq '' -or $SQLInstance -eq $null){
        $SQLInstance = Read-Host -Prompt 'Input your SQL server instance name'
        $SQLDatabase = Read-Host -Prompt 'Input the "Orchestrator" database name (leave empty if same name)'
    }
    Write-Host "Executing Remote SQL queries on $SQLDatabase database from $SQLInstance SQL Server instance..." -ForegroundColor Cyan
    runSQLQueries $SQLInstance $SQLDatabase
    Write-Host "Exporting Log Files on local server..." -ForegroundColor Cyan 
    exportLogFiles
	
	#region Waiting for background tasks to complete
	if (-not ($msinfo32.HasExited)) {
		#Write-Host "Waiting for System Information to complete in the background. Please wait...."
		Wait-Process -InputObject $msinfo32
	}
	#endregion
	
	
	Write-Output ""
	$completionDateTime = (Get-Date).ToString("yyyy-MM-dd__HH.mm.ss.fff")  
	Write-Host "Now compressing..." -ForegroundColor Yellow
	$script:SQLResultSetCounter = $null
	Stop-Transcript | out-null
	$ProgressPreference = 'Continue'

	$resultingZipFile_FullPath = (Join-Path -Path (Split-Path $MyInvocation.MyCommand.Definition) -ChildPath "$($resultPrefix)_$($RoleFoundAbbr)_$($resultDateTime).zip")
	if ( $PSVersionTable.PSVersion.Major -lt 4 ) { 
		Compress-ZipFile  $resultFolder.FullName $resultingZipFile_FullPath 
	}
	else {
		MakeNewZipFile $resultFolder.FullName $resultingZipFile_FullPath 
	}
	Remove-Item $resultFolder -Recurse 

	if ([Environment]::UserInteractive) {
		#CLS
		#Write-Host ""
		#Write-Host "Info about Secure File Exchange:"
		#Write-Host "https://docs.microsoft.com/en-US/troubleshoot/azure/general/secure-file-exchange-transfer-files"
		Write-Host ""
		Write-Host "Collection completed at $completionDateTime. (local time)"
		Write-Host -NoNewline "FINISHED! Please upload output "; Write-Host -NoNewline -ForegroundColor Yellow "$($resultPrefix)_$($RoleFoundAbbr)_$($resultDateTime).zip"; Write-Host -NoNewline " saved in folder "; Write-Host -NoNewline -ForegroundColor Yellow "$(Split-Path $MyInvocation.MyCommand.Definition)"; Write-Host -NoNewline " to case workspace"
		Write-Host ""
		Write-Host "Press ENTER to navigate to the resulting zip file..." -ForegroundColor Cyan
		Read-Host " "  
		  
		start (join-path $env:Windir explorer.exe) -ArgumentList "/select, ""$resultingZipFile_FullPath"""
	}
		
# SIG # Begin signature block
# MIInwQYJKoZIhvcNAQcCoIInsjCCJ64CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC2bHL6UmCSTauS
# 4QxB148rXO1wWbrYpda4xHTUAp+jOKCCDXYwggX0MIID3KADAgECAhMzAAACy7d1
# OfsCcUI2AAAAAALLMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjIwNTEyMjA0NTU5WhcNMjMwNTExMjA0NTU5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC3sN0WcdGpGXPZIb5iNfFB0xZ8rnJvYnxD6Uf2BHXglpbTEfoe+mO//oLWkRxA
# wppditsSVOD0oglKbtnh9Wp2DARLcxbGaW4YanOWSB1LyLRpHnnQ5POlh2U5trg4
# 3gQjvlNZlQB3lL+zrPtbNvMA7E0Wkmo+Z6YFnsf7aek+KGzaGboAeFO4uKZjQXY5
# RmMzE70Bwaz7hvA05jDURdRKH0i/1yK96TDuP7JyRFLOvA3UXNWz00R9w7ppMDcN
# lXtrmbPigv3xE9FfpfmJRtiOZQKd73K72Wujmj6/Su3+DBTpOq7NgdntW2lJfX3X
# a6oe4F9Pk9xRhkwHsk7Ju9E/AgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUrg/nt/gj+BBLd1jZWYhok7v5/w4w
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzQ3MDUyODAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAJL5t6pVjIRlQ8j4dAFJ
# ZnMke3rRHeQDOPFxswM47HRvgQa2E1jea2aYiMk1WmdqWnYw1bal4IzRlSVf4czf
# zx2vjOIOiaGllW2ByHkfKApngOzJmAQ8F15xSHPRvNMmvpC3PFLvKMf3y5SyPJxh
# 922TTq0q5epJv1SgZDWlUlHL/Ex1nX8kzBRhHvc6D6F5la+oAO4A3o/ZC05OOgm4
# EJxZP9MqUi5iid2dw4Jg/HvtDpCcLj1GLIhCDaebKegajCJlMhhxnDXrGFLJfX8j
# 7k7LUvrZDsQniJZ3D66K+3SZTLhvwK7dMGVFuUUJUfDifrlCTjKG9mxsPDllfyck
# 4zGnRZv8Jw9RgE1zAghnU14L0vVUNOzi/4bE7wIsiRyIcCcVoXRneBA3n/frLXvd
# jDsbb2lpGu78+s1zbO5N0bhHWq4j5WMutrspBxEhqG2PSBjC5Ypi+jhtfu3+x76N
# mBvsyKuxx9+Hm/ALnlzKxr4KyMR3/z4IRMzA1QyppNk65Ui+jB14g+w4vole33M1
# pVqVckrmSebUkmjnCshCiH12IFgHZF7gRwE4YZrJ7QjxZeoZqHaKsQLRMp653beB
# fHfeva9zJPhBSdVcCW7x9q0c2HVPLJHX9YCUU714I+qtLpDGrdbZxD9mikPqL/To
# /1lDZ0ch8FtePhME7houuoPcMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGaEwghmdAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAALLt3U5+wJxQjYAAAAAAsswDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIGkIqx148HfkQr6cpEaUgO+G
# BWoARvgHQ0vR50LtrINxMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQBBktAWZqV4Ej3nHO1Exek1DpWh86npTvixhNbWf9AzaJ8mdpmBz9//
# n7Ge6bVFi2OyNXKoG3gEeeSqp0pFCVfJRyGsY8uTwLNZNS7G76pT9bYkYtSijbfc
# UfEgJxUvZklotU1Aqva1ykVBzdv5vkUtyAcujiUVw663rIhFEEsfDzHXrtvcnCwW
# 00Zhew9M9agKrkTJVOmoCIcA4oo7b4wFZj4sC/iIqvDPeE0f+zuMpgUNOju02ngH
# Wgq6ntHGNJbS7D5VhFIuH2OGEY0HG/jPH9HK2GjCHKwwPl9NlVt9PHXdQVg2ggip
# 7KKamIbSAC6Pqi48sJCVGD9K3yFggLw+oYIXKTCCFyUGCisGAQQBgjcDAwExghcV
# MIIXEQYJKoZIhvcNAQcCoIIXAjCCFv4CAQMxDzANBglghkgBZQMEAgEFADCCAVkG
# CyqGSIb3DQEJEAEEoIIBSASCAUQwggFAAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIEz0ERH9nna7Rreso6mVtrUa2VgS34EEpbJOtihtWiw/AgZjT994
# 6QEYEzIwMjIxMDI4MTAwMDA1LjE0M1owBIACAfSggdikgdUwgdIxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBF
# U046MDg0Mi00QkU2LUMyOUExJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2WgghF4MIIHJzCCBQ+gAwIBAgITMwAAAbJuQAN/bqmUkgABAAABsjAN
# BgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0y
# MjA5MjAyMDIyMDFaFw0yMzEyMTQyMDIyMDFaMIHSMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBP
# cGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjA4NDIt
# NEJFNi1DMjlBMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAyqJlMh17+VDisL4GaXl/
# 9a6r/EpPGt9sbbceh+ZD6pkA3gbI7vc8XfL04B+m3tB/aNyV1Y4ZQH4fMG7CWVjI
# /d/HgxjzO+4C4HfsW+jK2c0LYMqdWtWUc5VwZQv0KeaEM0wDb+eySMh/YiiIb0nS
# otivx268d1An0uLY+r2C7JJv2a9QvrSiCyUI72CSHoWIQPAyvBSvxaNrqMWlROfL
# y2DQ3RycI3bDh8qSnmplxtRgViJwtJv/oDukcK1frGeOrCGYmiJve+QonJXFu4Ut
# GFVfEf3lvQsd42GJ+feO+jaP7/hBXXSMSldVb6IL0GxO1Hr3G9ONTnVmA/sFHhgM
# RarsmzKVI6/kHlMdMNdF/XzhRHMWFPJvw5lApjuaoyHtzwnzDWwQzhcNQXZRk3Lz
# b01ULMba190RdlofEXxGbGlBgHHKFnBjWui24hL6B83Z6r6GQBPeKkafz8qYPAO3
# MBud+5eMCmB5mrCBxgnykMn7L/FTqi7MnPUG97lNOKGSIDvBCxB7pHrRmT10903P
# DQwrmeJHO5BkC3gYj3oWGOGVRZxRk4KS/8lcz84a7+uBKmVjB2Y8vPN8O1fK7L8Y
# JTkjiXTyDqKJ9fKkyChiSRx44ADPi/HXHQE6dlZ8jd9LCo1S+g3udxNP4wHhWm9/
# VAGmmMEBBS6+6Lp4IbQwJU0CAwEAAaOCAUkwggFFMB0GA1UdDgQWBBSZ8ieAXNkR
# mU+SMM5WW4FIMNpqcTAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBf
# BgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmww
# bAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0El
# MjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUF
# BwMIMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEA3Ee27cXMhpto
# NtaqzB0oGUCEpdEI37kJIyK/ZNhriLZC5Yib732mLACEOEAN9uqivXPIuL3ljoZC
# e8hZSB14LugvVm1nJ73bNgr4Qh/BhmaFL4IfiKd8DNS+xwdkXfCWslR89QgMZU/S
# UJhWx72aC68bR2qRjhrJA8Qc68m5uBllo52D83x0id3p8Z45z7QOgbMH4uJ45snZ
# DQC0S3dc3eJfwKnr51lNfzHAT8u+FHA+lv/6cqyE7tNW696fB1PCoH8tPoI09oSX
# AV4rEqupFM8xsd6D6L4qcEt/CaERewyDazVBfskjF+9P3qZ3R6IyOIwQ7bYts7OY
# sw13csg2jACdEEAm1f7f97f3QH2wwYwen5rVX6GCzrYCikGXSn/TSWLfQM3nARDk
# h/flmTtv9PqkTHqslQNgK2LvMJuKSMpNqcGc5z33MYyV6Plf58L+TkTFQKs6zf9X
# MZEJm3ku9VBJ1aqr9AzNMSaKbixvMBIr2KYSSM21lnK8LUKxRwPW+gWS2V3iYoyM
# T64MRXch10P4OtGT3idXM09K5ld7B9U6dcdJ6obvEzdXt+XZovi/U6Evb4nA7VPH
# cHSKs7U72ps10mTfnlue13VFJUqAzbYoUEeegvsmzulGEGJoqZVNAag5v6PVBrur
# 5yLEajjxWH2TfkEOwlL8MuhcVI8OXiYwggdxMIIFWaADAgECAhMzAAAAFcXna54C
# m0mZAAAAAAAVMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZp
# Y2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMy
# MjVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEA5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51
# yMo1V/YBf2xK4OK9uT4XYDP/XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY
# 6GB9alKDRLemjkZrBxTzxXb1hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9
# cmmvHaus9ja+NSZk2pg7uhp7M62AW36MEBydUv626GIl3GoPz130/o5Tz9bshVZN
# 7928jaTjkY+yOSxRnOlwaQ3KNi1wjjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDua
# Rr3tpK56KTesy+uDRedGbsoy1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74
# kpEeHT39IM9zfUGaRnXNxF803RKJ1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2
# K26oElHovwUDo9Fzpk03dJQcNIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5
# TI4CvEJoLhDqhFFG4tG9ahhaYQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZk
# i1ugpoMhXV8wdJGUlNi5UPkLiWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9Q
# BXpsxREdcu+N+VLEhReTwDwV2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3Pmri
# Lq0CAwEAAaOCAd0wggHZMBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUC
# BBYEFCqnUv5kxJq+gpE8RjUpzxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJl
# pxtTNRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9y
# eS5odG0wEwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUA
# YgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU
# 1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2Ny
# bC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIw
# MTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0w
# Ni0yMy5jcnQwDQYJKoZIhvcNAQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/yp
# b+pcFLY+TkdkeLEGk5c9MTO1OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulm
# ZzpTTd2YurYeeNg2LpypglYAA7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM
# 9W0jVOR4U3UkV7ndn/OOPcbzaN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECW
# OKz3+SmJw7wXsFSFQrP8DJ6LGYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4
# FOmRsqlb30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3Uw
# xTSwethQ/gpY3UA8x1RtnWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPX
# fx5bRAGOWhmRaw2fpCjcZxkoJLo4S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVX
# VAmxaQFEfnyhYWxz/gq77EFmPWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGC
# onsXHRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU
# 5nR0W2rRnj7tfqAxM328y+l7vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEG
# ahC0HVUzWLOhcGbyoYIC1DCCAj0CAQEwggEAoYHYpIHVMIHSMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJl
# bGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNO
# OjA4NDItNEJFNi1DMjlBMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQCOEn4R7JJF+fYoI2yOf1wX0BRJOqCBgzCB
# gKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBBQUA
# AgUA5wWSbzAiGA8yMDIyMTAyODA3Mjc0M1oYDzIwMjIxMDI5MDcyNzQzWjB0MDoG
# CisGAQQBhFkKBAExLDAqMAoCBQDnBZJvAgEAMAcCAQACAheuMAcCAQACAhFZMAoC
# BQDnBuPvAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEA
# AgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQEFBQADgYEAFWLUFxTwyZ2/0KDE
# DNgNIQwdRox1G4WXuleP9gO+PtWYEpLMkYvLGzVQGmVXXp2NPuUoFFj3hqNxkQQP
# M45wvajMRNKrEp4FWEqS2+HOEMeFZZz5lg0GurVAQfesEAyttSW+PB03zZtVBJoz
# DZRJP3P/yzx7ovJxiF2OS1NibzoxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBQQ0EgMjAxMAITMwAAAbJuQAN/bqmUkgABAAABsjANBglghkgBZQME
# AgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJ
# BDEiBCAM4XvtMS7xMkHRBFYc2Cy48sGCqop/podl7foCm/n/hjCB+gYLKoZIhvcN
# AQkQAi8xgeowgecwgeQwgb0EIFN4zjzn4T63g8RWJ5SgUpfs9XIuj+fO76G0k8Ib
# Tj41MIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAGy
# bkADf26plJIAAQAAAbIwIgQgJbeT+2xQVEJ7O8Ent0IMJu4dtKfUvbq9U0WrZWDb
# evkwDQYJKoZIhvcNAQELBQAEggIAQoBg1G3pBFcFawxGyuZKLnF++qOQqHuFUFQD
# 0OiHy5Jhy3hSwey0B/eU/RS99+bdTWushtGzGT9K3gM5N5sWyhQK4pOcqNZTqhqs
# ZOqcusVQ5TaoEurks4NvUi1WRYTM8pnKUu4pFmJOhr3p2oVnIGXnEny903ISh4Z2
# XlN0seNW1ifg9juyHfqnL1cXu9Pj+IK8JLfLbG5izcyj+U+Vo/iXw2X9tOHW/e3Y
# oZIN64j4YZtEPJ81lV05yGCxI0lIVRZQVlcZ3dlFWBEpTD/k3jJ5x9yibHJyKwT8
# /zYhNGGrUf9fCmFATVpJvurC6DCU3CCZVCYkKYIWhoN9yIB7aslUn8dhMtzOhGPs
# Kn61Mfg6ZRI2emOrOV2l6G9mqSo4+d3S2IDtDphVsc95r41wxP32bKN5SRiHWTtC
# aMFzVJCJNVUIwfxnsUEsQ6RKQ8rSrhXgOLUH/xyhCD0wPlq7k5fdEeippXb8Je4k
# pWcgVXCzWiwn71qyNlkbqeZW1kbdD+QfKWXWfn/VySZM9M8R/nU2iDn3JRzXfhu7
# tnW1wytOCJeMPGOFnxVh3ofnDKS8e+88YDFhMatji44boWMWUe/cFOR9yzbpgnVJ
# RkjoAgWy7cDZvT8ARiZQJObB1/p+963gDJJnSAcmEYSZGgZ78J8hMX/NE6a5jErA
# s6aDcnI=
# SIG # End signature block

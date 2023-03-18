<#
	.SYNOPSIS
		Sets an icon for a folder
	
	.DESCRIPTION
		Set-FolderIcon function sets an icon for a folder. It looks for .exe or .ico files and creates a desktop.ini file with the required parameters.
	
	.PARAMETER Path
		Folder path
	
	.PARAMETER ICO_priority
		Sets the priority for .ico files:
		- icon.ico: icon.ico takes priority over .exe files and other .ico files.
		- any: any .ico file takes priority over .exe files.
		- like_folder: an .ico file with the same name as the folder takes priority over .exe files and other .ico files.
	
	.PARAMETER Filter
		Enter the names of folders you want to ignore in the format 'FIRST_NAME' OR 'FIRST_NAME', 'SECOND_NAME' OR @('FIRST_NAME', 'SECOND_NAME').
	
	.PARAMETER Single
		This is a switch. Use it if you want to add an icon only for the selected folder and not for child folders.

	.PARAMETER Dependencies
		You can use own rules for icon search in the hashtable @{"Inno Setup 5" = "Compil32.exe"; "DaVinci Resolve" = "Resolve.exe"}

	.PARAMETER folders_Path
		Enter the path to folders in the format 'FIRST_PATH' OR @('FIRST_PATH', 'SECOND_PATH') OR 'FIRST_PATH', 'SECOND_PATH'. You can also use the pipeline, e.g. "'FIRST_PATH', 'SECOND_PATH' | Set-FolderIcon".
		If the folders_Path parameter is not set, the system FolderDialog is opened.
	
	.EXAMPLE
		PS C:\> Set-FolderIcon -folders_Path @('C:\test', 'D:\test')
		The script sets icons for the 'C:\test' and 'D:\test' folders.
	
	.EXAMPLE
		PS C:\> Set-FolderIcon -folders_Path 'C:\test' -ICO_priority icon.ico -Filter 'ProcessHacker'
		The icon.ico file takes priority over .exe files and other .ico files.
	
	.EXAMPLE
		PS C:\> 'C:\test' | Set-FolderIcon -ICO_priority icon.ico -Filter 'OneCommander', 'Windows Kits'
		Example how to use the pipeline to get the path for the folders_Path parameter.
	
	.EXAMPLE
		PS C:\> Set-FolderIcon -ICO_priority icon.ico -Filter @('OneCommander', 'Windows Kits')
		Example how to use system FolderDialog to get the path for the folders_Path parameter.
	
	.NOTES
		Written by JustDj (justdj.ca@gmail.com)
		- The format for the folders_Path parameter is mandatory: 'C:\test' OR @('C:\test', 'D:\test') OR 'C:\test', 'D:\test'.
		- The format for the Filter parameter is mandatory: 'test' OR @('test', 'test2') OR 'test', 'test2'.
		- If the folders_Path parameter is not set, the system FolderDialog is opened.
		- You can use the pipeline for the folders_Path parameter.
	
	.LINK
		https://github.com/JustDj-Git/Set-FolderIcon
#>
function Set-FolderIcon {
	param
	(
		[Parameter(ValueFromPipeline = $true)]
		[Alias("folderPath")]
		[string[]]$Path,
		[ValidateSet('any', 'icon.ico', 'like_folder')]
		[Alias("Priority")]
		$ICO_priority,
		[Alias("folder_filter")]
		[string[]]$Filter,
		[Alias("alone")]
		[switch]$Single,
		$Dependencies
	)
	
begin {
		#avoid Path var
		$scriptPath = $Path
		$ErrorActionPreference = 'Stop'
		$pattern_regex = '[^\w\s]'
		$pattern_regex_symbols = '[!#$%&()+;@^_{}~№]'
		$pattern_regex_digits = '\d+'
		$timer = [Diagnostics.Stopwatch]::StartNew()
	}
	process {
		try {
			if (!($scriptPath)) {
				Add-Type -AssemblyName System.Windows.Forms
				$OpenFolderDialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
				$OpenFolderDialog.Description = 'Select a folder'
				$OpenFolderDialog.rootfolder = 'MyComputer'
				
				if (($OpenFolderDialog.ShowDialog() -eq 'OK')) {
					$scriptPath = $OpenFolderDialog.SelectedPath
				} else {
					Write-Host 'Do it again' -ForegroundColor Red
					return
				}
			}
			
			foreach ($i in $scriptPath) {
				if (Test-Path -Path $i) {
					Write-Host "Folder `'$i`' exist" -ForegroundColor Green
					if ($Single) {
						$folders += Get-Item -Path $i -ErrorAction SilentlyContinue
					} else {
						$folders += Get-ChildItem -Path $i -Directory -ErrorAction SilentlyContinue
					}
				} else {
					Write-Host "Folder `'$i`' does not exist" -ForegroundColor Red
					continue
				}
			}
			
			#main foreach
			foreach ($folder in $folders) {
				if ($Filter -notcontains $($folder.Name)) {
					
					#delete desktop.ini's
					Get-ChildItem -Path "$($folder.FullName)" -Filter "desktop.ini" -Hidden -Recurse -Depth 1 | Remove-Item -Force
					
					#Null elements
					$array_string = @()
					$exeFiles = ''
					$exeFile_name_checked = @()
					$array_exes = @()
					
					#setting proper attributes
					$folder.Attributes = 'Directory', 'ReadOnly'
					
					#sometimes need lowercase
					[string]$name_folder = ($folder.Name).ToLower()
					
					#pattern for regex. Delete all symbols + numbers
					$name_folder = $name_folder -replace $pattern_regex, '' -replace "$pattern_regex_digits", '' -replace "$pattern_regex_symbols", ''
					
					#full path to folder
					[string]$full_path_folder = $folder.FullName
					
					#Find in dependencies
					$LastDirName = Split-Path -Path $full_path_folder -Leaf
					if (($Dependencies) -and ($Dependencies.ContainsKey($LastDirName))){
						$value = $Dependencies[$LastDirName]
						if (Test-Path $full_path_folder\$value) {
							$exeFiles = Get-ChildItem -Path $full_path_folder -Filter $value
						}
					}
					
					if (!($exeFiles)) {
						######## check each name + with a space
						$tmpr = $name_folder.Trim().Split(' ')
						foreach ($j in $tmpr) {
							$array_string += $j
						} #foreach
						
						# Loop through each element of an array
						foreach ($item in $array_string) {
							# looking for folders in which there is a file with a name corresponding to an array element
							$files = Get-ChildItem -Path $full_path_folder -Recurse -Filter "*$item*.exe"
							
							# if the file exists, add the element to the array of found elements
							if ($files) {
								$exeFile_name_checked += $item
							}
						} #foreach
						
						#take the first object, resolve the path
						$get_first = $exeFile_name_checked[0]
						$exeFiles = Get-ChildItem -Path $full_path_folder -Filter "*$get_first*.exe" -Recurse | Select-Object -First 1
						########
					}

					#if not found - ICO
					if ((!($exeFiles)) -or ($ICO_priority -eq 'like_folder')) {
						$exeFiles = Get-ChildItem -Path $full_path_folder -Filter "$name_folder.ico" -Recurse | Select-Object -First 1
					} elseif ((!($exeFiles)) -or ($ICO_priority -eq 'icon.ico')) {
						$exeFiles = Get-ChildItem -Path $full_path_folder -Filter 'icon.ico' | Select-Object -First 1
					} elseif ((!($exeFiles)) -or ($ICO_priority -eq 'any')) {
						$exeFiles = Get-ChildItem -Path $full_path_folder -Filter '*.ico' -Recurse | Select-Object -First 1
					}
					
					#if not found at all - assign any
					if (!($exeFiles)) {
						$exeFiles = (Get-ChildItem -Path $full_path_folder -Filter '*.exe' -Recurse)
						
						foreach ($exeFile in $exeFiles) {
							[string]$name_exe = ($exeFile.BaseName).ToLower() -replace "$pattern_regex_symbols", '' -replace "$pattern_regex", '' -replace "$pattern_regex_digits", ''
							if (($name_folder -like $name_exe) `
								-or ($name_folder -match $name_exe) `
								-or ($name_folder.Contains($name_exe)) `
								-or ($name_exe.Contains($name_folder))) {
								
								$array_exes += $exeFile
							}
						}
						#take the first found
						$exeFiles = $array_exes[0]
					} ########if not found at all - assign any
					
					#found
					if ($exeFiles) {
						$first_part = ''
						[string]$name_exe = ($exeFiles.BaseName).ToLower()
						
						Write-Verbose -Message "`n$name_exe"
						
						#if exe is nested somewhere
						if (($exeFiles.DirectoryName -ne $full_path_folder)) {
							
							#splitting the full path into an array
							$exe_array = ($exeFiles.DirectoryName).Split('\')
							$folder_array = ($full_path_folder).Split('\')
							
							#Compare 2 arrays, write down the difference
							$diff = (Compare-Object -ReferenceObject $exe_array -DifferenceObject $folder_array).InputObject
							
							#Difference
							foreach ($k in $diff) {
								$first_part = $first_part + '\' + $k
							}
						} ###found
						
						# make a temporary folder with desktop.ini
						$tmpDir = (Join-Path -Path $env:TEMP -ChildPath ([IO.Path]::GetRandomFileName()))
						$null = mkdir -Path $tmpDir -Force
						$tmp = "$tmpDir\desktop.ini"
						
						#
						if ($first_part) {
							$value = '.' + "$first_part\$exeFiles" + ',0'
						} else {
							$value = '.\' + $exeFiles + ',0'
						}
						
						$ini = @"
[.ShellClassInfo]
IconResource=$value
InfoTip=$exeFiles
[ViewState]
Mode=
Vid=
FolderType=Generic
"@
						
						$null = New-Item -Path $tmp -Value $ini
						
						(Get-Item -Path $tmp).Attributes = 'Archive, System, Hidden'
						
						$shell = New-Object -ComObject Shell.Application
						$shell.NameSpace($full_path_folder).MoveHere($tmp, 0x0004 + 0x0010 + 0x0400)
						# FOF_SILENT         0x0004 don't display progress UI
						# FOF_NOCONFIRMATION 0x0010 don't display confirmation UI, assume "yes" 
						# FOF_NOERRORUI      0x0400 don't put up error UI
						
						Remove-Item -Path $tmpDir -Force
						
						Write-Host "`n$($exeFiles.Name) ==> '$($folder.Name)' folder" -ForegroundColor Green
						Write-Verbose -Message "`n$value"
						
					} else {
						write-host "`nProper exe not found in $full_path_folder" -ForegroundColor Yellow
					}
				} else {
					write-host "`nFolder $folder filtered" -ForegroundColor Yellow
				}
				##TODO counter
			} #main foreach
		} catch {
			Write-Host "`n$_" -ForegroundColor Red
			Write-Host "`n$($_.ScriptStackTrace)`n" -ForegroundColor Red
		}
	}
	end {
		$timer.Stop()
		$timeRound = [Math]::Round(($timer.Elapsed.TotalSeconds), 2)
		$timer.Reset()
		Write-Host "`nTask completed in $timeRound`s" -ForegroundColor Cyan
	}
}
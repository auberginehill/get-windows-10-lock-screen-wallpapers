<#
Get-Windows10LockScreenWallpapers.ps1


# Credit:               robledosm: "Save lockscreen as wallpaper"
# Forked from:          robledosm: "Save lockscreen as wallpaper"
# Original version:     robledosm: "Save lockscreen as wallpaper": https://github.com/robledosm/save-lockscreen-as-wallpaper
#>


[CmdletBinding()]
param (
    [Parameter(ValueFromPipeline=$true,
    ValueFromPipelineByPropertyName=$true,
 #  Mandatory=$true,
      HelpMessage="`r`nOutput: In which folder or directory would you like to find the new wallpapers? `r`n`r`nPlease enter a valid file system path to a directory (a full path name of a folder such as C:\Windows). `r`n`r`nNotes:`r`n`t- If the path name includes space characters, please enclose the path in quotation marks (single or double). `r`n`t- To exit this script, please press [Ctrl] + C`r`n")]
    [Alias("Path","Destination","OutputFolder")]
    [string]$Output = "$($env:USERPROFILE)\Pictures\Wallpapers",
    [ValidateScript({
            # Credit: Mike F Robbins: "PowerShell Advanced Functions: Can we build them better?" http://mikefrobbins.com/2015/03/31/powershell-advanced-functions-can-we-build-them-better-with-parameter-validation-yes-we-can/
            If ($_ -match '^(?!^(PRN|AUX|CLOCK\$|NUL|CON|COM\d|LPT\d|\..*)(\..+)?$)[^\x00-\x1f\\?*:\"";|/]+$') {
                $True
            }
            Else {
                Throw "`r`n - $_ is either not a valid subfolder name or it is not recommended. `r`n - If the subfolder name includes space characters, please enclose it in quotation marks."
            }
            })]
    [Parameter(HelpMessage="`r`nSubfolder: In which subfolder or subdirectory under the directory defined with the -Output parameter would you like to find the new portrait (vertical) pictures? `r`n`r`nPlease enter a name for the folder (i.e. a 'DirectoryName' value, which doesn't include a path, and doesn't reveal any parent or root directories). `r`n`r`nNotes:`r`n`t- If the name includes space characters, please enclose it in quotation marks (single or double). `r`n")]
    [Alias("SubfolderForThePortraitPictures","SubfolderForTheVerticalPictures","SubfolderName")]
    [string]$Subfolder = "Vertical",
    [switch]$Force,
    [Alias("NoPortrait","NoSubfolder","Exclude")]
    [switch]$ExcludePortrait,
    [Alias("IncludeCurrentLockScreenBackgroundHive")]
    [switch]$Include,
    [switch]$Log,
    [switch]$Open,
    [switch]$Audio
)


Begin {

    # Set some common variables
    $ErrorActionPreference = "Stop"
    $computer = $env:COMPUTERNAME
    $log_filename = "spotlight_log.csv"
    $date = Get-Date -Format g
    $existing_landscape = @{}
    $existing_portrait = @{}
    $supplementary_log = @{}
    $existing_hashes = @{}
    $new_landscape = @()
    $unique_files = @()
    $new_portrait = @()
    $source_paths = @()
    $horizontal = @()
    $new_files = @()
    $vertical = @()
    $results = @()
    $remarks = @()
    $new_images = 0
    $num_vertical = 0
    $num_horizontal = 0
    $text_verbose = "Please consider checking that the destination folder '$Output', where the new wallpapers are ought to be written (and which is set with the -Output parameter), was typed correctly, and that it is a valid file system path, which points to a directory. If the path name includes space characters, please enclose the whole path string after the -Output parameter in quotation marks (single or double)."
    $empty_line = ""


    # Function used to convert bytes to MB or GB or TB                                                      # Credit: clayman2: "Disk Space"
    function ConvertBytes {
        Param (
            $size
        )
        If ($size -eq $null) {
            [string]'-'
        } ElseIf ($size -eq 0) {
            [string]'-'
        } ElseIf ($size -lt 1MB) {
            $file_size = $size / 1KB
            $file_size = [Math]::Round($file_size, 0)
            [string]$file_size + ' KB'
        } ElseIf ($size -lt 1GB) {
            $file_size = $size / 1MB
            $file_size = [Math]::Round($file_size, 1)
            [string]$file_size + ' MB'
        } ElseIf ($size -lt 1TB) {
            $file_size = $size / 1GB
            $file_size = [Math]::Round($file_size, 1)
            [string]$file_size + ' GB'
        } Else {
            $file_size = $size / 1TB
            $file_size = [Math]::Round($file_size, 1)
            [string]$file_size + ' TB'
        } # else
    } # function (ConvertBytes)


    # Define the Windows Spotlight lock screen wallpapers' source path
    # Default: "$($env:USERPROFILE)\AppData\Local\Packages\Microsoft.Windows.ContentDeliveryManager*\LocalState\Assets"
    # Default: "$($env:LOCALAPPDATA)\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets"
    # Source: http://www.askvg.com/windows-10-wallpapers-and-lock-screen-backgrounds/
    # Source: https://www.cnet.com/how-to/where-to-find-the-windows-spotlight-photos/
    # Source: http://www.windowscentral.com/how-save-all-windows-spotlight-lockscreen-images
    # Source: https://github.com/hashhar/Windows-Hacks/tree/master/scheduled-tasks/save-windows-spotlight-lockscreens
    # Source: https://answers.microsoft.com/en-us/insider/forum/insider_wintp-insider_personal/windows-10041-windows-spotlight-lock-screen/5b1cddaf-7057-443b-99b6-8c3486a75262


        # Exit earlier than Windows 10 machines
        If (([System.Environment]::OSVersion.Version).Major -lt 10) {

                    $os = ((Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer).Caption).TrimEnd()

                    If ($os.Contains(","))      { $os = $os.Replace(",","") }
                    If ($os.Contains("(R)"))    { $os = $os.Replace("(R)","") }

                    $empty_line | Out-String
                    $version_text = "Windows Spotlight lock screen is available in all desktop editions of Windows 10. $os doesn't contain Windows Spotlight."
                    Write-Output $version_text
                    $empty_line | Out-String
                    Exit

        } Else {

            # (Source 1)
            # Read the registry
            # Source: http://www.winhelponline.com/blog/find-file-name-lock-screen-image-current-displayed/
            # Source: https://www.tenforums.com/tutorials/38717-windows-spotlight-background-images-find-save-windows-10-a.html
            $reg_key = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Lock Screen\Creative"

                            If ((Test-Path $reg_key) -eq $false) {
                                $continue = $true
                            } Else {
                                $registry = Get-ItemProperty -Path $reg_key | Select-Object -ExpandProperty LandscapeAssetPath
                                # PortraitAssetPath...
                            } # Else (If Test-Path $reg_key)

            If (($registry -eq $null) -or ($registry -eq "") -or ($registry -eq " ") -or ($registry -eq "0")) {

                # (Source 2)
                # Estimate the probable source path
                # Note: The test-procedure will fail, if there is other than one directory inside the $($env:USERPROFILE)\AppData\Local\Packages\ folder, which begins with "Microsoft.Windows.ContentDelivery".
                # Note: Alternative and equally working $source_paths variable value is listed below for skipping the test procedure.
                # $source_paths = "$($env:LOCALAPPDATA)\Packages\Microsoft.Windows.ContentDeliveryManager*\LocalState\Assets"
                $primary_candidate = "$($env:LOCALAPPDATA)\Packages\Microsoft.Windows.ContentDeliveryManager*\LocalState\Assets"
                $root_candidate = "$($env:LOCALAPPDATA)\Packages\Microsoft.Windows.ContentDelivery*"


                    $test_1 = Get-ChildItem -Path "$root_candidate" -Directory -Force -ErrorAction SilentlyContinue
                    If (($test_1.Count) -eq 1) {
                        $root_folder = $test_1.FullName
                        $source_path_candidate = "$root_folder\LocalState\Assets"
                    } Else {
                        $continue = $true
                    } # Else (If $test_1.Count)

                    $test_2 = Test-Path $source_path_candidate -ErrorAction SilentlyContinue
                    If (($test_2) -eq $true) {
                        $source_paths += "$source_path_candidate\*"
                        $source_two = $true
                    } Else {
                        $continue = $true
                    } # Else (If $test_2 -eq $true)


                    If (($test_2 -eq $false) -or ($test_2 -eq $null) -or ($test_2 -eq "")) {

                        # (Source 3)
                        # Resort to retrieving the wallpapers from the current lock screen hive
                        # Note: The wallpapers may have at least .png and .jpg file extensions.
                        # Source: http://www.ohmancorp.com/RefWin-windows8-change-pre-login-screen-background.asp
                        # Source: https://social.technet.microsoft.com/Forums/windows/en-US/a8db890c-204f-404a-bf74-3aa4c895b183/cant-customize-lock-or-logon-screen-backgrounds?forum=W8ITProPreRel
                        # Credit: klumsy: "Call Windows Runtime Classes from PowerShell": http://stackoverflow.com/questions/14115415/call-windows-runtime-classes-from-powershell
                        [Windows.System.UserProfile.LockScreen, Windows.System.UserProfile, ContentType=WindowsRuntime] | Out-Null
                        $current_lockscreen_background_full_path = ([Windows.System.UserProfile.LockScreen]::OriginalImageFile).LocalPath
                        $source_path_runner_up = ([System.IO.Path]::GetDirectoryName($current_lockscreen_background_full_path))
                        $source_paths += "$source_path_runner_up\*"
                        $source_three = $true
                    } Else {
                        $continue = $true
                    } # Else (If $test_2 -eq $false)


                    If (($source_paths -eq $null) -or ($source_paths -eq "") -or ($source_paths -eq " ") -or ($source_paths -eq "0")) {

                        # No Source: Display an error message in console and exit
                        $empty_line | Out-String
                        Write-Warning "The Windows SpotLight lock screen wallpapers' source '$primary_candidate' doesn't seem to exist."
                        $empty_line | Out-String
                        $location_text = "Didn't detect the Windows Spotlight lock screen backgrounds' default source path '$primary_candidate'. Is the Spotlight feature enabled? For further info, please visit https://technet.microsoft.com/en-us/itpro/windows/manage/windows-spotlight, http://www.windowscentral.com/how-enable-windows-spotlight or http://www.winhelponline.com/blog/find-file-name-lock-screen-image-current-displayed/ `r`nExit Code 1: Couldn't locate the source folder."
                        Write-Output $location_text
                        $empty_line | Out-String
                        Exit

                    } Else {
                        $continue = $true
                    } # Else (If $source_paths -eq $null)
            } Else {
                $source_paths += "$registry\*"
                $source_one = $true
            } # Else (If $registry -eq $null)
        } # Else (If [System.Environment]::OSVersion)


    # Add the path of the current lock screen wallpaper hive to the sources, if set to do so with the -Include parameter
    If (($Include) -and (($source_three) -ne $true)) {

        # Note: The wallpapers may have at least .png and .jpg file extensions.
        # Source: http://www.ohmancorp.com/RefWin-windows8-change-pre-login-screen-background.asp
        # Source: https://social.technet.microsoft.com/Forums/windows/en-US/a8db890c-204f-404a-bf74-3aa4c895b183/cant-customize-lock-or-logon-screen-backgrounds?forum=W8ITProPreRel
        # Credit: klumsy: "Call Windows Runtime Classes from PowerShell": http://stackoverflow.com/questions/14115415/call-windows-runtime-classes-from-powershell
        [Windows.System.UserProfile.LockScreen, Windows.System.UserProfile, ContentType=WindowsRuntime] | Out-Null
        $current_lockscreen_background_full_path = ([Windows.System.UserProfile.LockScreen]::OriginalImageFile).LocalPath
        $source_path_secondary = ([System.IO.Path]::GetDirectoryName($current_lockscreen_background_full_path))
        $source_paths += "$source_path_secondary\*"

    } ElseIf (($Include) -and (($source_three) -eq $true)) {
        $continue = $true
    } Else {
        $continue = $true
    } # Else (If $Include)


    # Select files greater than 200kb                                                                       # Credit: robledosm: "Save lockscreen as wallpaper"
    $sources = Get-ChildItem -Path $source_paths -Force -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 200kb }

    If ($sources.Count -ge 1) {

        # Test if the Output-path ("Destination") exists
        If ((Test-Path $Output -PathType Leaf) -eq $true) {

            # File: Display an error message in console and exit
            $empty_line | Out-String
            Write-Warning "-Output: '$Output' seems to point to a file."
            $empty_line | Out-String
            Write-Verbose $text_verbose -verbose
            $empty_line | Out-String
            $exit_text = "Couldn't open the -Output folder '$Output' since it's a file."
            Write-Output $exit_text
            $empty_line | Out-String
            Exit

        } ElseIf ((Test-Path $Output) -eq $false) {

            If ($Force) {

                # If the Force was used, create the destination folder ($Output)
                New-Item "$Output" -ItemType Directory -Force | Out-Null
                $continue = $true

            } Else {

                # No Destination: Display an error message in console
                $empty_line | Out-String
                Write-Warning "'$Output' doesn't seem to exist."
                $empty_line | Out-String
                Write-Verbose $text_verbose -verbose

                # Offer the user an option to create the defined $Output ("Destination") folder             # Credit: lamaar75: "Creating a Menu": http://powershell.com/cs/forums/t/9685.aspx
                # Source: "Adding a Simple Menu to a Windows PowerShell Script": https://technet.microsoft.com/en-us/library/ff730939.aspx
                $title_corner = "Create folder '$Output' with this script?"
                $message = " "

                $yes = New-Object System.Management.Automation.Host.ChoiceDescription    "&Yes",    "Yes:     tries to create a new folder and to copy the found new wallpapers there."
                $no = New-Object System.Management.Automation.Host.ChoiceDescription     "&No",     "No:      exits from this script (similar to Ctrl + C)."
                $exit = New-Object System.Management.Automation.Host.ChoiceDescription   "&Exit",   "Exit:    exits from this script (similar to Ctrl + C)."
                $abort = New-Object System.Management.Automation.Host.ChoiceDescription  "&Abort",  "Abort:   exits from this script (similar to Ctrl + C)."
                $cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel", "Cancel:  exits from this script (similar to Ctrl + C)."

                $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $exit, $abort, $cancel)
                $choice_result = $host.ui.PromptForChoice($title_corner, $message, $options, 1)

                    switch ($choice_result)
                        {
                            0 {
                            "Yes. Creating the folder.";
                            New-Item "$Output" -ItemType Directory -Force | Out-Null
                            $continue = $true
                            }
                            1 {
                            $empty_line | Out-String
                            "No. Exiting from Spotlight lock screen background retrieval script.";
                            $empty_line | Out-String
                            Exit
                            }
                            2 {
                            $empty_line | Out-String
                            "Exit. Exiting from Spotlight lock screen background retrieval script.";
                            $empty_line | Out-String
                            Exit
                            }
                            3 {
                            $empty_line | Out-String
                            "Abort. Exiting from Spotlight lock screen background retrieval script.";
                            $empty_line | Out-String
                            Exit
                            }
                            4 {
                            $empty_line | Out-String
                            "Cancel. Exiting from Spotlight lock screen background retrieval script.";
                            $empty_line | Out-String
                            Exit
                            } # 4
                        } # switch
            } # Else If $Force)
        } Else {
            $continue = $true
        } # Else (Test-Path $Output -PathType Leaf)

                # Resolve the Output-path ("Destination") (if the Output-path is specified as relative) and remove the last character if it's \
                # Source: https://technet.microsoft.com/en-us/library/ee692804.aspx
                # Source: http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string#32608908
                $real_output_path = (Resolve-Path -Path $Output).Path
                If ((($real_output_path).EndsWith("\")) -eq $true) { $real_output_path = $real_output_path -replace ".{1}$" }
    } Else {
        $empty_line | Out-String
        $exit_text = "Didn't find any files at the source folder '$($source_paths -join ', ')' (Exit 2)."
        Write-Output $exit_text
        $empty_line | Out-String
        Exit

    } # Else (If $sources.Count)

                                            If ($ExcludePortrait) {
                                                $continue = $true
                                            } Else {
                                                # Create a subfolder for the portrait (vertical) pictures if it doesn't exist
                                                $subfolder_path = "$real_output_path\$Subfolder"
                                                If ((Test-Path $subfolder_path) -eq $true) {
                                                    $continue = $true
                                                } Else {
                                                    New-Item "$subfolder_path" -ItemType Directory -Force | Out-Null
                                                } # Else (If Test-Path $subfolder_path)
                                            } # Else (If $ExcludePortrait)

        # Try to process each file (-gt 200kb) only once
        If ($Include) {

            If (($source_one) -eq $true) {

                $source_one_hashes = Get-FileHash -Path "$registry\*" -Algorithm SHA256
                ForEach ($entity in $source_one_hashes) {
                    $supplementary_log.Add("$($entity.Path)", "$($entity.Hash)")
                } # ForEach $entity)

            } ElseIf (($source_two) -eq $true) {

                $source_two_hashes = Get-FileHash -Path "$source_path_candidate\*" -Algorithm SHA256
                ForEach ($entity in $source_two_hashes) {
                    $supplementary_log.Add("$($entity.Path)", "$($entity.Hash)")
                } # ForEach $entity)

            } Else {
                $continue = $true
            } # Else (If $source_one)


                        $secondary_repo = ([System.IO.Path]::GetDirectoryName($current_lockscreen_background_full_path))
                        $potential_duplicates = Get-ChildItem -Path "$secondary_repo\*" -Force -ErrorAction SilentlyContinue

                        ForEach ($file_item in $potential_duplicates) {
                                $hash_value = Get-FileHash -Path $file_item.FullName -Algorithm SHA256 | Select-Object -ExpandProperty Hash
                                    If ($supplementary_log.ContainsValue("$hash_value")) {

                                        # Exclude the duplicate source item and return to the top of the program loop (ForEach $file_item)
                                        $empty_line | Out-String
                                        Write-Warning "Duplicate image found."
                                        $empty_line | Out-String
                                        $duplicate_text = "Skipping '$file_item' from the files to be processed."
                                        Write-Output $duplicate_text
                                        Continue

                                    } Else {
                                        $supplementary_log.Add("$($file_item.FullName)", "$($hash_value)")
                                    } # Else (If $supplementary_log.ContainsValue)
                        } # ForEach ($file_item)

            # Select files greater than 200kb                                                           # Credit: robledosm: "Save lockscreen as wallpaper"
            ForEach ($sample in $supplementary_log.Keys) {
                $unique_files += Get-ChildItem -Path $sample -Force -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 200kb }
            } # ForEach $sample)
        } Else {
            $unique_files = $sources
        } # Else (If $Include)


                    # Compare the hash values
                    # Note: Probably requires at least PowerShell version 4
                    # Source: https://technet.microsoft.com/en-us/library/ee692803.aspx
                    # Source: https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.utility/get-filehash
                    $landscape_hashes = Get-FileHash -Path "$real_output_path\*" -Algorithm SHA256
                    ForEach ($item in $landscape_hashes) {
                        $existing_landscape.Add("$($item.Path)", "$($item.Hash)")
                        $existing_hashes.Add("$($item.Path)", "$($item.Hash)")
                    } # ForEach $item)
                    $portrait_hashes = Get-FileHash -Path "$subfolder_path\*" -Algorithm SHA256
                    ForEach ($entity in $portrait_hashes) {
                        $existing_portrait.Add("$($entity.Path)", "$($entity.Hash)")
                        $existing_hashes.Add("$($entity.Path)", "$($entity.Hash)")
                    } # ForEach $entity)
                    ForEach ($file_candidate in $unique_files) {
                        $new_hash = Get-FileHash -Path $file_candidate.FullName -Algorithm SHA256 | Select-Object -ExpandProperty Hash
                            If ($existing_hashes.ContainsValue("$new_hash")) {
                                $continue = $true
                            } Else {
                                $new_files += $file_candidate
                            } # Else (If $existing_hashes.ContainsValue)
                    } # ForEach ($file_candidate)


        If (($new_files.Count) -ge 1) {
            $continue = $true
        } Else {
            $empty_line | Out-String
            $text = "Didn't find any new images."
            Write-Output $text
            $empty_line | Out-String
            Exit

        } # Else (If $existing_hashes.ContainsValue)


                    # Set Image Swithes (used in Step 2)
                    # Source: http://nicholasarmstrong.com/2010/02/exif-quick-reference/
                    # Source: http://msdn.microsoft.com/en-us/library/ms630826%28v=vs.85%29.aspx
                    # Credit: Franck Richard: "Use PowerShell to Remove Metadata and Resize Images": http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html
                    $contrast   = @{ 0 = "Normal"; 1 = "Low"; 2 = "High" }
                    $custrender = @{ 0 = "Normal"; 1 = "Custom" }
                    $exposure   = @{ 0 = "Auto"; 1 = "Manual"; 2 = "Auto Bracket" }
                    $filesrc    = @{ 1 = "Film scanner"; 2 = "Reflection print scanner"; 3 = "Digital camera" }
                    $flash      = @{ 0 = "No Flash"; 10 = "Flash off"; 1 = "Flash on"; 11 = "Flash auto" }
                    $focal      = @{ 1 = "None"; 2 = "Inches"; 3 = "Centimetres" ; 4 = "Millimetres"; 5 = "Micrometres" }
                    $gain       = @{ 0 = "None"; 1 = "Low gain up"; 2 = "High gain up"; 3 = "Low gain down"; 4 = "High gain down" }
                    $metering   = @{ 0 = "Unknown"; 1 = "Average"; 2 = "Center-weighted average" ; 3 = "Spot"; 4 = "Multi-spot"; 5 = "Multi-segment"; 6 = "Partial"; 255 = "Unknown" }
                    $orient     = @{ 1 = "Horizontal"; 3 = "Rotate 180 degrees"; 6 = "Rotate 90 degrees clockwise" ; 8 = "Rotate 270 degrees clockwise" }
                    $saturation = @{ 0 = "Normal"; 1 = "Low"; 2 = "High" }
                    $scene      = @{ 0 = "Standard"; 1 = "Landscape"; 2 = "Portrait"; 3 = "Night" }
                    $sdr        = @{ 0 = "Unknown"; 1 = "Macro"; 2 = "Close" ; 3 = "Distant" }
                    $sensing    = @{ 1 = "Not defined"; 2 = "One-chip colour area"; 3 = "Two-chip colour area" ; 4 = "Three-chip colour area"; 5 = "Colour sequential area"; 7 = "Trilinear"; 8 = "Colour sequential linear" }
                    $sharpness  = @{ 0 = "Normal"; 1 = "Soft"; 2 = "Hard" }
                    $unit       = @{ 1 = "None"; 2 = "Inches"; 3 = "Centimetres" }
                    $white      = @{ 0 = "Auto"; 1 = "Manual" }


        # Test if the Microsoft Windows Image Acquisition (WIA) service is enabled (used in Step 1)
        $test_wia = Get-Service | where { $_.ServiceName -eq "stisvc" } -ErrorAction SilentlyContinue
        $wia_startup_type = (Get-WmiObject -Class Win32_Service -ComputerName $computer -Filter "Name='stisvc'").StartMode

        If (($test_wia -eq $null) -or ($wia_startup_type -eq 'Disabled')) {

            # If the WIA service is not enabled, display an error message in console and exit
            $empty_line | Out-String
            Write-Warning "The Microsoft Windows Image Acquisition (WIA) service 'stisvc' doesn't seem to be enabled."
            $empty_line | Out-String
            Write-Verbose "The WIA service is needed for opening and saving the the image files. For futher instructions, how to enable the WIA service, please for example see http://kb.winzip.com/kb/entry/207/ " -verbose
            $empty_line | Out-String
            $exit_text = "Didn't open the $($new_files.Count) found new files for further examination."
            Write-Output $exit_text
            $empty_line | Out-String
            Exit

        } Else {
            $continue = $true
        } # Else (If $test_wia)

                    # Set the progress bar variables ($id denominates different progress bars, if more than one is being displayed)
                    $activity           = "Processing $($new_files.Count) Spotlight lock screen backgrounds"
                    $status             = " "
                    $task               = "Setting Initial Variables"
                    $operations         = (($new_files.Count) + 2)
                    $total_steps        = (($new_files.Count) + 3)
                    $task_number        = 0
                    $id                 = 1

                    # Start the progress bar if there is more than one unique file to process
                    If (($new_files.Count) -ge 2) {
                        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete ((0.000002 / $total_steps) * 100)
                    } # If ($new_files.Count)

} # Begin




Process {

    # Process each file (-gt 200kb)
    $image = New-Object -ComObject WIA.ImageFile
    ForEach ($file in $new_files) {

                    # Increment the step counter
                    $task_number++

                    # Update the progress bar if there is more than one unique file to process
                    $activity = "Processing $($new_files.Count) Spotlight lock screen backgrounds - Step $task_number/$operations"
                    If (($new_files.Count) -ge 2) {
                        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $file.Name -PercentComplete (($task_number / $total_steps) * 100)
                    } # If ($new_files.Count)

        # Step 1
        # Load the image as an ImageFile COM object with Microsoft Windows Image Acquisition (WIA)
        # Note: To edit the image: WIA.ImageProcess... $ip = New-Object -ComObject WIA.ImageProcess  ...    $ip.FilterInfos | fl *
        # Note: For data retriaval also: System.Drawing.Image.PropertyItems
        # Note: In the default installations of Windows Server 2003 and 2012 the Windows Image Acquisition (WIA) service is not enabled by default.
        # Note: On Windows Server 2008 the WIA service is not installed by default. To add this feature, please see the WinZip Knowledgebase link below.
        # Source: http://kb.winzip.com/kb/entry/207/
        # Source: https://msdn.microsoft.com/en-us/library/windows/desktop/ms630506(v=vs.85).aspx
        # Source: https://blogs.msdn.microsoft.com/powershell/2009/03/30/image-manipulation-in-powershell/
        # Source: http://stackoverflow.com/questions/4304821/get-startup-type-of-windows-service-using-powershell
        # Source: https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.management/get-wmiobject
        # Source: https://social.microsoft.com/Forums/en-US/4dfe4eec-2b9b-4e6e-a49e-96f5a108c1c8/using-powershell-as-a-photoshop-replacement?forum=Offtopic
        # $disabled_wia = Get-WmiObject -Class Win32_Service -ComputerName $computer -Filter "Name='stisvc'" | Where-Object { $_.StartMode -eq 'Disabled' }
        # $test_wia = Get-Service | where {$_.DisplayName -like "*Windows Image Acquisition*" } -ErrorAction SilentlyContinue
        $image.LoadFile($file.FullName)

            # Source: https://msdn.microsoft.com/en-us/library/system.io.path_methods(v=vs.110).aspx
            # $destination_path = "$real_output_path\$($file.BaseName).jpg"
            # $new_filename = [System.IO.Path]::ChangeExtension($file.Name, "jpg")
            $new_filename = [System.IO.Path]::ChangeExtension($file.Name, $image.FileExtension)

            # Determine the picture orientation
            If (($image.Width -ge "1600") -and ($image.Width -gt $image.Height)) {
                $type = "Landscape"
                $orientation = "Horizontal"
                $horizontal += "$($file.FullName)"
                $destination_path = "$real_output_path\$new_filename"
            } ElseIf (($image.Height -ge "1200") -and ($image.Height -gt $image.Width)) {
                $type = "Portrait"
                $orientation = "Vertical"
                $vertical += "$($file.FullName)"
                $destination_path = "$real_output_path\$Subfolder\$new_filename"
            } Else {
                $continue = $true
            } # Else (If $image.Width)


            # Step 2
            # Retrieve image properties (n ~300)
            # Source: https://msdn.microsoft.com/en-us/library/ms630826(VS.85).aspx#SharedSample012
            If ($image.IsIndexedPixelFormat -eq $true )     { $remarks += "Pixel data contains palette indexes" }                               Else { $continue = $true }
            If ($image.IsAlphaPixelFormat -eq $true )       { $remarks += "Pixel data has alpha information" }                                  Else { $continue = $true }
            If ($image.IsExtendedPixelFormat -eq $true )    { $remarks += "Pixel data has extended color information (16 bit/channel)" }        Else { $continue = $true }
            If ($image.IsAnimated -eq $true )               { $remarks += "Image is animated" }                                                 Else { $continue = $true }

            Try {
                $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256 | Select-Object -ExpandProperty Hash
            } Catch { Write-Debug $_.Exception }

            # Source: https://www.experts-exchange.com/questions/25100459/I-need-to-send-the-details-of-a-jpg-file-to-an-array-any-windows-api-to-do-this-or-get-me-started.html
            # Source: https://social.technet.microsoft.com/Forums/windowsserver/en-US/16124c53-4c7f-41f2-9a56-7808198e102a/attribute-seems-to-give-byte-array-how-to-convert-to-string?forum=winserverpowershell
            # Source: http://compgroups.net/comp.databases.ms-access/handy-routine-for-getting-file-metad/1484921
            If ($image.Properties.Exists('40091')) {
                $bytes_title = $image.Properties.Item('40091')
                $title = ($bytes_title.Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }             Else { $continue = $true }

            If ($image.Properties.Exists('40092')) {
                $bytes_comment = $image.Properties.Item('40092')
                $comment = ($bytes_comment.Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }         Else { $continue = $true }

            If ($image.Properties.Exists('40093')) {
                $bytes_author = $image.Properties.Item('40093')
                $author = ($bytes_author.Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }           Else { $continue = $true }

            If ($image.Properties.Exists('40094')) {
                $bytes_keywords = $image.Properties.Item('40094')
                $keywords = ($bytes_keywords.Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }       Else { $continue = $true }

            If ($image.Properties.Exists('40095')) {
                $bytes_subject = $image.Properties.Item('40095')
                $subject = ($bytes_subject.Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }         Else { $continue = $true }

            If ($image.Properties.Exists('36864')) {
                $ExifVer = (($image.Properties.Item('36864')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }              Else { $continue = $true }

            If ($image.Properties.Exists('40960')) {
                $FlashpixVersion = (($image.Properties.Item('40960')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }      Else { $continue = $true }

            If ($image.Properties.Exists('37510')) {
                $UserComment = (($image.Properties.Item('37510')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }          Else { $continue = $true }

        #    If ($image.Properties.Exists('37500')) {
        #        $MakerNote = (($image.Properties.Item('37500')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }            Else { $continue = $true }


            $results += $obj_file = New-Object -TypeName PSCustomObject -Property @{

                        'ActiveFrame'                   = $image.ActiveFrame
                        'BaseName'                      = $file.BaseName
                        'Created'                       = $file.CreationTime
                        'Destination'                   = $destination_path
                        'Directory'                     = $file.Directory
                        'Original File Extension'       = $file.Extension
                        'File Name Without Extension'   = $file.BaseName
                        'File Original'                 = $file.Name
                        'File New'                      = $new_filename
                        'File'                          = $new_filename
                        'FileExtension'                 = $image.FileExtension
                        'FormatID'                      = $image.FormatID
                        'FrameCount'                    = $image.FrameCount
                        'SHA256'                     	= $hash
                        'Height'                        = $image.Height
                        'Home'                          = $file.FullName
                        'HorizontalResolution'          = $image.HorizontalResolution
                        'IsAlphaPixelFormat'            = $image.IsAlphaPixelFormat
                        'IsAnimated'                    = $image.IsAnimated
                        'IsExtendedPixelFormat'         = $image.IsExtendedPixelFormat
                        'IsIndexedPixelFormat'          = $image.IsIndexedPixelFormat
                        'Modified'                      = $file.LastWriteTime
                        'Orientation'                   = $orientation
                        'PixelDepth'                    = $image.PixelDepth
                        'raw_size'                      = $file.Length
                        'Remarks'                       = ($remarks -join ', ')
                        'Saved to a New Location'       = $date
                        'Size'                     	    = (ConvertBytes ($file.Length))
                        'Source'                        = $file.FullName
                        'Type'                          = $type
                        'VerticalResolution'            = $image.VerticalResolution
                        'Width'                         = $image.Width


                        # Source: https://www.experts-exchange.com/questions/25100459/I-need-to-send-the-details-of-a-jpg-file-to-an-array-any-windows-api-to-do-this-or-get-me-started.html
                        'Title'                         = If ($image.Properties.Exists('40091')) { $title }                                                     Else {" "}
                        'Comment'                       = If ($image.Properties.Exists('40092')) { $comment }                                                   Else {" "}
                        'Author'                        = If ($image.Properties.Exists('40093')) { $author }                                                    Else {" "}
                        'Keywords'                      = If ($image.Properties.Exists('40094')) { $keywords }                                                  Else {" "}
                        'Subject'                       = If ($image.Properties.Exists('40095')) { $subject }                                                   Else {" "}


                        # Source: http://www.exiv2.org/tags.html
                        'ExifVersion'                   = If ($image.Properties.Exists('36864')) { $ExifVer }                                                   Else {" "}
                        'FlashpixVersion'               = If ($image.Properties.Exists('40960')) { $FlashpixVersion }                                           Else {" "}
                    #    'MakerNote'                     = If ($image.Properties.Exists('37500')) { $MakerNote }                                                 Else {" "}
                        'UserComment'                   = If ($image.Properties.Exists('37510')) { $UserComment }                                               Else {" "}
                        'ActiveArea'                    = If ($image.Properties.Exists('50829')) { $image.Properties.Item('50829').Value }                      Else {" "}
                        'AnalogBalance'                 = If ($image.Properties.Exists('50727')) { $image.Properties.Item('50727').Value }                      Else {" "}
                        'AntiAliasStrength'             = If ($image.Properties.Exists('50738')) { $image.Properties.Item('50738').Value }                      Else {" "}
                        'Artist'                        = If ($image.Properties.Exists('315'))   { $image.Properties.Item('315').Value }                        Else {" "}
                        'AsShotICCProfile'              = If ($image.Properties.Exists('50831')) { $image.Properties.Item('50831').Value }                      Else {" "}
                        'AsShotNeutral'                 = If ($image.Properties.Exists('50728')) { $image.Properties.Item('50728').Value }                      Else {" "}
                        'AsShotPreProfileMatrix'        = If ($image.Properties.Exists('50832')) { $image.Properties.Item('50832').Value }                      Else {" "}
                        'AsShotWhiteXY'                 = If ($image.Properties.Exists('50729')) { $image.Properties.Item('50729').Value }                      Else {" "}
                        'BaselineExposure'              = If ($image.Properties.Exists('50730')) { $image.Properties.Item('50730').Value }                      Else {" "}
                        'BaselineNoise'                 = If ($image.Properties.Exists('50731')) { $image.Properties.Item('50731').Value }                      Else {" "}
                        'BaselineSharpness'             = If ($image.Properties.Exists('50732')) { $image.Properties.Item('50732').Value }                      Else {" "}
                        'BatteryLevel'                  = If ($image.Properties.Exists('33423')) { $image.Properties.Item('33423').Value }                      Else {" "}
                        'BayerGreenSplit'               = If ($image.Properties.Exists('50733')) { $image.Properties.Item('50733').Value }                      Else {" "}
                        'BestQualityScale'              = If ($image.Properties.Exists('50780')) { $image.Properties.Item('50780').Value }                      Else {" "}
                        'BitsPerSample'                 = If ($image.Properties.Exists('258'))   { $image.Properties.Item('258').Value }                        Else {" "}
                        'BlackLevel'                    = If ($image.Properties.Exists('50714')) { $image.Properties.Item('50714').Value }                      Else {" "}
                        'BlackLevelDeltaH'              = If ($image.Properties.Exists('50715')) { $image.Properties.Item('50715').Value }                      Else {" "}
                        'BlackLevelDeltaV'              = If ($image.Properties.Exists('50716')) { $image.Properties.Item('50716').Value }                      Else {" "}
                        'BlackLevelRepeatDim'           = If ($image.Properties.Exists('50713')) { $image.Properties.Item('50713').Value }                      Else {" "}
                        'BodySerialNumber'              = If ($image.Properties.Exists('42033')) { $image.Properties.Item('42033').Value }                      Else {" "}
                        'BrightnessValue'               = If ($image.Properties.Exists('37379')) { $image.Properties.Item('37379').Value }                      Else {" "}
                        'Byte_AsShotProfileName'        = If ($image.Properties.Exists('50934')) { $image.Properties.Item('50934').Value }                      Else {" "}
                        'Byte_CameraCalibrationSignat'  = If ($image.Properties.Exists('50931')) { $image.Properties.Item('50931').Value }                      Else {" "}
                        'Byte_CFAPattern'               = If ($image.Properties.Exists('33422')) { $image.Properties.Item('33422').Value }                      Else {" "}
                        'Byte_CFAPlaneColor'            = If ($image.Properties.Exists('50710')) { $image.Properties.Item('50710').Value }                      Else {" "}
                        'Byte_ClipPath'                 = If ($image.Properties.Exists('343'))   { $image.Properties.Item('343').Value }                        Else {" "}
                        'Byte_DNGBackwardVersion'       = If ($image.Properties.Exists('50707')) { $image.Properties.Item('50707').Value }                      Else {" "}
                        'Byte_DNGPrivateData'           = If ($image.Properties.Exists('50740')) { $image.Properties.Item('50740').Value }                      Else {" "}
                        'Byte_DNGVersion'               = If ($image.Properties.Exists('50706')) { $image.Properties.Item('50706').Value }                      Else {" "}
                        'Byte_DotRange'                 = If ($image.Properties.Exists('336'))   { $image.Properties.Item('336').Value }                        Else {" "}
                        'Byte_GPSAltitudeRef'           = If ($image.Properties.Exists('5'))     { $image.Properties.Item('5').Value }                          Else {" "}
                        'Byte_GPSVersionID'             = If ($image.Properties.Exists('0'))     { $image.Properties.Item('0').Value }                          Else {" "}
                        'Byte_ImageResources'           = If ($image.Properties.Exists('34377')) { $image.Properties.Item('34377').Value }                      Else {" "}
                        'Byte_LocalizedCameraModel'     = If ($image.Properties.Exists('50709')) { $image.Properties.Item('50709').Value }                      Else {" "}
                        'Byte_OriginalRawFileName'      = If ($image.Properties.Exists('50827')) { $image.Properties.Item('50827').Value }                      Else {" "}
                        'Byte_PreviewApplicationName'   = If ($image.Properties.Exists('50966')) { $image.Properties.Item('50966').Value }                      Else {" "}
                        'Byte_PreviewApplicationVersion'= If ($image.Properties.Exists('50967')) { $image.Properties.Item('50967').Value }                      Else {" "}
                        'Byte_PreviewSettingsDigest'    = If ($image.Properties.Exists('50969')) { $image.Properties.Item('50969').Value }                      Else {" "}
                        'Byte_PreviewSettingsName'      = If ($image.Properties.Exists('50968')) { $image.Properties.Item('50968').Value }                      Else {" "}
                        'Byte_ProfileCalibrationSignat' = If ($image.Properties.Exists('50932')) { $image.Properties.Item('50932').Value }                      Else {" "}
                        'Byte_ProfileCopyright'         = If ($image.Properties.Exists('50942')) { $image.Properties.Item('50942').Value }                      Else {" "}
                        'Byte_ProfileName'              = If ($image.Properties.Exists('50936')) { $image.Properties.Item('50936').Value }                      Else {" "}
                        'Byte_RawDataUniqueID'          = If ($image.Properties.Exists('50781')) { $image.Properties.Item('50781').Value }                      Else {" "}
                        'Byte_TIFFEPStandardID'         = If ($image.Properties.Exists('37398')) { $image.Properties.Item('37398').Value }                      Else {" "}
                        'Byte_XMLPacket'                = If ($image.Properties.Exists('700'))   { $image.Properties.Item('700').Value }                        Else {" "}
                        'CalibrationIlluminant1'        = If ($image.Properties.Exists('50778')) { $image.Properties.Item('50778').Value }                      Else {" "}
                        'CalibrationIlluminant2'        = If ($image.Properties.Exists('50779')) { $image.Properties.Item('50779').Value }                      Else {" "}
                        'CameraCalibration1'            = If ($image.Properties.Exists('50723')) { $image.Properties.Item('50723').Value }                      Else {" "}
                        'CameraCalibration2'            = If ($image.Properties.Exists('50724')) { $image.Properties.Item('50724').Value }                      Else {" "}
                        'CameraOwnerName'               = If ($image.Properties.Exists('42032')) { $image.Properties.Item('42032').Value }                      Else {" "}
                        'CameraSerialNumber'            = If ($image.Properties.Exists('50735')) { $image.Properties.Item('50735').Value }                      Else {" "}
                        'CellLength'                    = If ($image.Properties.Exists('265'))   { $image.Properties.Item('265').Value }                        Else {" "}
                        'CellWidth'                     = If ($image.Properties.Exists('264'))   { $image.Properties.Item('264').Value }                        Else {" "}
                        'CFALayout'                     = If ($image.Properties.Exists('50711')) { $image.Properties.Item('50711').Value }                      Else {" "}
                        'CFAPattern'                    = If ($image.Properties.Exists('41730')) { $image.Properties.Item('41730').Value }                      Else {" "}
                        'CFARepeatPatternDim'           = If ($image.Properties.Exists('33421')) { $image.Properties.Item('33421').Value }                      Else {" "}
                        'ChromaBlurRadius'              = If ($image.Properties.Exists('50737')) { $image.Properties.Item('50737').Value }                      Else {" "}
                        'ColorimetricReference'         = If ($image.Properties.Exists('50879')) { $image.Properties.Item('50879').Value }                      Else {" "}
                        'ColorMap'                      = If ($image.Properties.Exists('320'))   { $image.Properties.Item('320').Value }                        Else {" "}
                        'ColorMatrix1'                  = If ($image.Properties.Exists('50721')) { $image.Properties.Item('50721').Value }                      Else {" "}
                        'ColorMatrix2'                  = If ($image.Properties.Exists('50722')) { $image.Properties.Item('50722').Value }                      Else {" "}
                        'ComponentsConfiguration'       = If ($image.Properties.Exists('37121')) { ($image.Properties.Item('37121').Value) -join (" ") }        Else {" "}
                        'CompressedBitsPerPixel'        = If ($image.Properties.Exists('37122')) { ($image.Properties.Item('37122').Value) | select -ExpandProperty Value } Else {" "}
                        'Compression'                   = If ($image.Properties.Exists('259'))   { $image.Properties.Item('259').Value }                        Else {" "}
                        'Copyright'                     = If ($image.Properties.Exists('33432')) { $image.Properties.Item('33432').Value }                      Else {" "}
                        'CurrentICCProfile'             = If ($image.Properties.Exists('50833')) { $image.Properties.Item('50833').Value }                      Else {" "}
                        'CurrentPreProfileMatrix'       = If ($image.Properties.Exists('50834')) { $image.Properties.Item('50834').Value }                      Else {" "}
                        'DefaultCropOrigin'             = If ($image.Properties.Exists('50719')) { $image.Properties.Item('50719').Value }                      Else {" "}
                        'DefaultCropSize'               = If ($image.Properties.Exists('50720')) { $image.Properties.Item('50720').Value }                      Else {" "}
                        'DefaultScale'                  = If ($image.Properties.Exists('50718')) { $image.Properties.Item('50718').Value }                      Else {" "}
                        'DeviceSettingDescription'      = If ($image.Properties.Exists('41995')) { $image.Properties.Item('41995').Value }                      Else {" "}
                        'DocumentName'                  = If ($image.Properties.Exists('269'))   { $image.Properties.Item('269').Value }                        Else {" "}
                        'ExifTag'                       = If ($image.Properties.Exists('34665')) { $image.Properties.Item('34665').Value }                      Else {" "}
                        'ExposureIndex'                 = If ($image.Properties.Exists('37397')) { $image.Properties.Item('37397').Value }                      Else {" "}
                        'ExposureIndex_2'               = If ($image.Properties.Exists('41493')) { $image.Properties.Item('41493').Value }                      Else {" "}
                        'ExposureProgram'               = If ($image.Properties.Exists('34850')) { $image.Properties.Item('34850').Value }                      Else {" "}
                        'ExtraSamples'                  = If ($image.Properties.Exists('338'))   { $image.Properties.Item('338').Value }                        Else {" "}
                        'FillOrder'                     = If ($image.Properties.Exists('266'))   { $image.Properties.Item('266').Value }                        Else {" "}
                        'FlashEnergy'                   = If ($image.Properties.Exists('37387')) { $image.Properties.Item('37387').Value }                      Else {" "}
                        'FlashEnergy_2'                 = If ($image.Properties.Exists('41483')) { $image.Properties.Item('41483').Value }                      Else {" "}
                        'FocalPlaneResolutionUnit'      = If ($image.Properties.Exists('37392')) { $image.Properties.Item('37392').Value }                      Else {" "}
                        'FocalPlaneXResolution'         = If ($image.Properties.Exists('37390')) { $image.Properties.Item('37390').Value }                      Else {" "}
                        'FocalPlaneYResolution'         = If ($image.Properties.Exists('37391')) { $image.Properties.Item('37391').Value }                      Else {" "}
                        'ForwardMatrix1'                = If ($image.Properties.Exists('50964')) { $image.Properties.Item('50964').Value }                      Else {" "}
                        'ForwardMatrix2'                = If ($image.Properties.Exists('50965')) { $image.Properties.Item('50965').Value }                      Else {" "}
                        'GPSAltitude'                   = If ($image.Properties.Exists('6'))     { ($image.Properties.Item('6').Value) | select -ExpandProperty Value } Else {" "}
                        'GPSAreaInformation'            = If ($image.Properties.Exists('28'))    { $image.Properties.Item('28').Value }                         Else {" "}
                        'GPSDateStamp'                  = If ($image.Properties.Exists('29'))    { $image.Properties.Item('29').Value }                         Else {" "}
                        'GPSDestBearing'                = If ($image.Properties.Exists('24'))    { $image.Properties.Item('24').Value }                         Else {" "}
                        'GPSDestBearingRef'             = If ($image.Properties.Exists('23'))    { $image.Properties.Item('23').Value }                         Else {" "}
                        'GPSDestDistance'               = If ($image.Properties.Exists('26'))    { $image.Properties.Item('26').Value }                         Else {" "}
                        'GPSDestDistanceRef'            = If ($image.Properties.Exists('25'))    { $image.Properties.Item('25').Value }                         Else {" "}
                        'GPSDestLatitude'               = If ($image.Properties.Exists('20'))    { $image.Properties.Item('20').Value }                         Else {" "}
                        'GPSDestLatitudeRef'            = If ($image.Properties.Exists('19'))    { $image.Properties.Item('19').Value }                         Else {" "}
                        'GPSDestLongitude'              = If ($image.Properties.Exists('22'))    { $image.Properties.Item('22').Value }                         Else {" "}
                        'GPSDestLongitudeRef'           = If ($image.Properties.Exists('21'))    { $image.Properties.Item('21').Value }                         Else {" "}
                        'GPSDifferential'               = If ($image.Properties.Exists('30'))    { $image.Properties.Item('30').Value }                         Else {" "}
                        'GPSImgDirection'               = If ($image.Properties.Exists('17'))    { ($image.Properties.Item('17').Value) | select -ExpandProperty Value } Else {" "}
                        'GPSImgDirectionRef'            = If ($image.Properties.Exists('16'))    { $image.Properties.Item('16').Value }                         Else {" "}
                        'GPSLatitude'                   = If ($image.Properties.Exists('2'))     { (($image.Properties.Item('2').Value) | select -ExpandProperty Value) -join (" ") } Else {" "}
                        'GPSLatitudeRef'                = If ($image.Properties.Exists('1'))     { $image.Properties.Item('1').Value }                          Else {" "}
                        'GPSLongitude'                  = If ($image.Properties.Exists('4'))     { (($image.Properties.Item('4').Value) | select -ExpandProperty Value) -join (" ") } Else {" "}
                        'GPSLongitudeRef'               = If ($image.Properties.Exists('3'))     { $image.Properties.Item('3').Value }                          Else {" "}
                        'GPSMapDatum'                   = If ($image.Properties.Exists('18'))    { $image.Properties.Item('18').Value }                         Else {" "}
                        'GPSMeasureMode'                = If ($image.Properties.Exists('10'))    { $image.Properties.Item('10').Value }                         Else {" "}
                        'GPSProcessingMethod'           = If ($image.Properties.Exists('27'))    { $image.Properties.Item('27').Value }                         Else {" "}
                        'GPSSatellites'                 = If ($image.Properties.Exists('8'))     { $image.Properties.Item('8').Value }                          Else {" "}
                        'GPSSpeed'                      = If ($image.Properties.Exists('13'))    { $image.Properties.Item('13').Value }                         Else {" "}
                        'GPSSpeedRef'                   = If ($image.Properties.Exists('12'))    { $image.Properties.Item('12').Value }                         Else {" "}
                        'GPSStatus'                     = If ($image.Properties.Exists('9'))     { $image.Properties.Item('9').Value }                          Else {" "}
                        'GPSTag'                        = If ($image.Properties.Exists('34853')) { $image.Properties.Item('34853').Value }                      Else {" "}
                        'GPSTimeStamp'                  = If ($image.Properties.Exists('7'))     { (($image.Properties.Item('7').Value) | select -ExpandProperty Value) -join (':') } Else {" "}
                        'GPSTrack'                      = If ($image.Properties.Exists('15'))    { $image.Properties.Item('15').Value }                         Else {" "}
                        'GPSTrackRef'                   = If ($image.Properties.Exists('14'))    { $image.Properties.Item('14').Value }                         Else {" "}
                        'GrayResponseCurve'             = If ($image.Properties.Exists('291'))   { $image.Properties.Item('291').Value }                        Else {" "}
                        'GrayResponseUnit'              = If ($image.Properties.Exists('290'))   { $image.Properties.Item('290').Value }                        Else {" "}
                        'HalftoneHints'                 = If ($image.Properties.Exists('321'))   { $image.Properties.Item('321').Value }                        Else {" "}
                        'HostComputer'                  = If ($image.Properties.Exists('316'))   { $image.Properties.Item('316').Value }                        Else {" "}
                        'ImageDescription'              = If ($image.Properties.Exists('270'))   { $image.Properties.Item('270').Value }                        Else {" "}
                        'ImageHistory'                  = If ($image.Properties.Exists('37395')) { $image.Properties.Item('37395').Value }                      Else {" "}
                        'ImageID'                       = If ($image.Properties.Exists('32781')) { $image.Properties.Item('32781').Value }                      Else {" "}
                        'ImageLength'                   = If ($image.Properties.Exists('257'))   { $image.Properties.Item('257').Value }                        Else {" "}
                        'ImageNumber'                   = If ($image.Properties.Exists('37393')) { $image.Properties.Item('37393').Value }                      Else {" "}
                        'ImageUniqueID'                 = If ($image.Properties.Exists('42016')) { $image.Properties.Item('42016').Value }                      Else {" "}
                        'ImageWidth'                    = If ($image.Properties.Exists('256'))   { $image.Properties.Item('256').Value }                        Else {" "}
                        'Indexed'                       = If ($image.Properties.Exists('346'))   { $image.Properties.Item('346').Value }                        Else {" "}
                        'InkNames'                      = If ($image.Properties.Exists('333'))   { $image.Properties.Item('333').Value }                        Else {" "}
                        'InkSet'                        = If ($image.Properties.Exists('332'))   { $image.Properties.Item('332').Value }                        Else {" "}
                        'Interlace'                     = If ($image.Properties.Exists('34857')) { $image.Properties.Item('34857').Value }                      Else {" "}
                        'InteroperabilityTag'           = If ($image.Properties.Exists('40965')) { $image.Properties.Item('40965').Value }                      Else {" "}
                        'IPTCNAA'                       = If ($image.Properties.Exists('33723')) { $image.Properties.Item('33723').Value }                      Else {" "}
                        'ISOSpeed'                      = If ($image.Properties.Exists('34867')) { $image.Properties.Item('34867').Value }                      Else {" "}
                        'ISOSpeedLatitudeyyy'           = If ($image.Properties.Exists('34868')) { $image.Properties.Item('34868').Value }                      Else {" "}
                        'ISOSpeedLatitudezzz'           = If ($image.Properties.Exists('34869')) { $image.Properties.Item('34869').Value }                      Else {" "}
                        'JPEGACTables'                  = If ($image.Properties.Exists('521'))   { $image.Properties.Item('521').Value }                        Else {" "}
                        'JPEGDCTables'                  = If ($image.Properties.Exists('520'))   { $image.Properties.Item('520').Value }                        Else {" "}
                        'JPEGInterchangeFormat'         = If ($image.Properties.Exists('513'))   { $image.Properties.Item('513').Value }                        Else {" "}
                        'JPEGInterchangeFormatLength'   = If ($image.Properties.Exists('514'))   { $image.Properties.Item('514').Value }                        Else {" "}
                        'JPEGLosslessPredictors'        = If ($image.Properties.Exists('517'))   { $image.Properties.Item('517').Value }                        Else {" "}
                        'JPEGPointTransforms'           = If ($image.Properties.Exists('518'))   { $image.Properties.Item('518').Value }                        Else {" "}
                        'JPEGProc'                      = If ($image.Properties.Exists('512'))   { $image.Properties.Item('512').Value }                        Else {" "}
                        'JPEGQTables'                   = If ($image.Properties.Exists('519'))   { $image.Properties.Item('519').Value }                        Else {" "}
                        'JPEGRestartInterval'           = If ($image.Properties.Exists('515'))   { $image.Properties.Item('515').Value }                        Else {" "}
                        'JPEGTables'                    = If ($image.Properties.Exists('347'))   { $image.Properties.Item('347').Value }                        Else {" "}
                        'LensInfo'                      = If ($image.Properties.Exists('50736')) { $image.Properties.Item('50736').Value }                      Else {" "}
                        'LensMake'                      = If ($image.Properties.Exists('42035')) { $image.Properties.Item('42035').Value }                      Else {" "}
                        'LensModel'                     = If ($image.Properties.Exists('42036')) { $image.Properties.Item('42036').Value }                      Else {" "}
                        'LensSerialNumber'              = If ($image.Properties.Exists('42037')) { $image.Properties.Item('42037').Value }                      Else {" "}
                        'LensSpecification'             = If ($image.Properties.Exists('42034')) { $image.Properties.Item('42034').Value }                      Else {" "}
                        'LightSource'                   = If ($image.Properties.Exists('37384')) { $image.Properties.Item('37384').Value }                      Else {" "}
                        'LinearizationTable'            = If ($image.Properties.Exists('50712')) { $image.Properties.Item('50712').Value }                      Else {" "}
                        'LinearResponseLimit'           = If ($image.Properties.Exists('50734')) { $image.Properties.Item('50734').Value }                      Else {" "}
                        'MakerNoteSafety'               = If ($image.Properties.Exists('50741')) { $image.Properties.Item('50741').Value }                      Else {" "}
                        'MaskedAreas'                   = If ($image.Properties.Exists('50830')) { $image.Properties.Item('50830').Value }                      Else {" "}
                        'NewSubfileType'                = If ($image.Properties.Exists('254'))   { $image.Properties.Item('254').Value }                        Else {" "}
                        'Noise'                         = If ($image.Properties.Exists('37389')) { $image.Properties.Item('37389').Value }                      Else {" "}
                        'NoiseProfile'                  = If ($image.Properties.Exists('51041')) { $image.Properties.Item('51041').Value }                      Else {" "}
                        'NoiseReductionApplied'         = If ($image.Properties.Exists('50935')) { $image.Properties.Item('50935').Value }                      Else {" "}
                        'NumberOfInks'                  = If ($image.Properties.Exists('334'))   { $image.Properties.Item('334').Value }                        Else {" "}
                        'OECF'                          = If ($image.Properties.Exists('34856')) { $image.Properties.Item('34856').Value }                      Else {" "}
                        'OpcodeList1'                   = If ($image.Properties.Exists('51008')) { $image.Properties.Item('51008').Value }                      Else {" "}
                        'OpcodeList2'                   = If ($image.Properties.Exists('51009')) { $image.Properties.Item('51009').Value }                      Else {" "}
                        'OpcodeList3'                   = If ($image.Properties.Exists('51022')) { $image.Properties.Item('51022').Value }                      Else {" "}
                        'OPIProxy'                      = If ($image.Properties.Exists('351'))   { $image.Properties.Item('351').Value }                        Else {" "}
                        'OriginalRawFileData'           = If ($image.Properties.Exists('50828')) { $image.Properties.Item('50828').Value }                      Else {" "}
                        'OriginalRawFileDigest'         = If ($image.Properties.Exists('50973')) { $image.Properties.Item('50973').Value }                      Else {" "}
                        'PageNumber'                    = If ($image.Properties.Exists('297'))   { $image.Properties.Item('297').Value }                        Else {" "}
                        'PhotometricInterpretation'     = If ($image.Properties.Exists('262'))   { $image.Properties.Item('262').Value }                        Else {" "}
                        'PlanarConfiguration'           = If ($image.Properties.Exists('284'))   { $image.Properties.Item('284').Value }                        Else {" "}
                        'Predictor'                     = If ($image.Properties.Exists('317'))   { $image.Properties.Item('317').Value }                        Else {" "}
                        'PreviewColorSpace'             = If ($image.Properties.Exists('50970')) { $image.Properties.Item('50970').Value }                      Else {" "}
                        'PreviewDateTime'               = If ($image.Properties.Exists('50971')) { $image.Properties.Item('50971').Value }                      Else {" "}
                        'PrimaryChromaticities'         = If ($image.Properties.Exists('319'))   { (($image.Properties.Item('319').Value) | select -ExpandProperty Value) -join (" ") } Else {" "}
                        'PrintImageMatching'            = If ($image.Properties.Exists('50341')) { $image.Properties.Item('50341').Value }                      Else {" "}
                        'ProcessingSoftware'            = If ($image.Properties.Exists('11'))    { ($image.Properties.Item('11').Value) | select -ExpandProperty Value } Else {" "}
                        'ProfileEmbedPolicy'            = If ($image.Properties.Exists('50941')) { $image.Properties.Item('50941').Value }                      Else {" "}
                        'ProfileHueSatMapData1'         = If ($image.Properties.Exists('50938')) { $image.Properties.Item('50938').Value }                      Else {" "}
                        'ProfileHueSatMapData2'         = If ($image.Properties.Exists('50939')) { $image.Properties.Item('50939').Value }                      Else {" "}
                        'ProfileHueSatMapDims'          = If ($image.Properties.Exists('50937')) { $image.Properties.Item('50937').Value }                      Else {" "}
                        'ProfileLookTableData'          = If ($image.Properties.Exists('50982')) { $image.Properties.Item('50982').Value }                      Else {" "}
                        'ProfileLookTableDims'          = If ($image.Properties.Exists('50981')) { $image.Properties.Item('50981').Value }                      Else {" "}
                        'ProfileToneCurve'              = If ($image.Properties.Exists('50940')) { $image.Properties.Item('50940').Value }                      Else {" "}
                        'Rating'                        = If ($image.Properties.Exists('18246')) { $image.Properties.Item('18246').Value }                      Else {" "}
                        'RatingPercent'                 = If ($image.Properties.Exists('18249')) { $image.Properties.Item('18249').Value }                      Else {" "}
                        'RawImageDigest'                = If ($image.Properties.Exists('50972')) { $image.Properties.Item('50972').Value }                      Else {" "}
                        'RecommendedExposureIndex'      = If ($image.Properties.Exists('34866')) { $image.Properties.Item('34866').Value }                      Else {" "}
                        'ReductionMatrix1'              = If ($image.Properties.Exists('50725')) { $image.Properties.Item('50725').Value }                      Else {" "}
                        'ReductionMatrix2'              = If ($image.Properties.Exists('50726')) { $image.Properties.Item('50726').Value }                      Else {" "}
                        'ReferenceBlackWhite'           = If ($image.Properties.Exists('532'))   { $image.Properties.Item('532').Value }                        Else {" "}
                        'RelatedSoundFile'              = If ($image.Properties.Exists('40964')) { $image.Properties.Item('40964').Value }                      Else {" "}
                        'RowInterleaveFactor'           = If ($image.Properties.Exists('50975')) { $image.Properties.Item('50975').Value }                      Else {" "}
                        'RowsPerStrip'                  = If ($image.Properties.Exists('278'))   { $image.Properties.Item('278').Value }                        Else {" "}
                        'SampleFormat'                  = If ($image.Properties.Exists('339'))   { $image.Properties.Item('339').Value }                        Else {" "}
                        'SamplesPerPixel'               = If ($image.Properties.Exists('277'))   { $image.Properties.Item('277').Value }                        Else {" "}
                        'SceneType'                     = If ($image.Properties.Exists('41729')) { $image.Properties.Item('41729').Value }                      Else {" "}
                        'SecurityClassification'        = If ($image.Properties.Exists('37394')) { $image.Properties.Item('37394').Value }                      Else {" "}
                        'SelfTimerMode'                 = If ($image.Properties.Exists('34859')) { $image.Properties.Item('34859').Value }                      Else {" "}
                        'SensingMethod'                 = If ($image.Properties.Exists('37399')) { $image.Properties.Item('37399').Value }                      Else {" "}
                        'SensitivityType'               = If ($image.Properties.Exists('34864')) { $image.Properties.Item('34864').Value }                      Else {" "}
                        'ShadowScale'                   = If ($image.Properties.Exists('50739')) { $image.Properties.Item('50739').Value }                      Else {" "}
                        'SMaxSampleValue'               = If ($image.Properties.Exists('341'))   { $image.Properties.Item('341').Value }                        Else {" "}
                        'SMinSampleValue'               = If ($image.Properties.Exists('340'))   { $image.Properties.Item('340').Value }                        Else {" "}
                        'Software'                      = If ($image.Properties.Exists('305'))   { $image.Properties.Item('305').Value }                        Else {" "}
                        'SpatialFrequencyResponse'      = If ($image.Properties.Exists('37388')) { $image.Properties.Item('37388').Value }                      Else {" "}
                        'SpatialFrequencyResponse_2'    = If ($image.Properties.Exists('41484')) { $image.Properties.Item('41484').Value }                      Else {" "}
                        'SpectralSensitivity'           = If ($image.Properties.Exists('34852')) { $image.Properties.Item('34852').Value }                      Else {" "}
                        'StandardOutputSensitivity'     = If ($image.Properties.Exists('34865')) { $image.Properties.Item('34865').Value }                      Else {" "}
                        'StripByteCounts'               = If ($image.Properties.Exists('279'))   { $image.Properties.Item('279').Value }                        Else {" "}
                        'StripOffsets'                  = If ($image.Properties.Exists('273'))   { $image.Properties.Item('273').Value }                        Else {" "}
                        'SubfileType'                   = If ($image.Properties.Exists('255'))   { $image.Properties.Item('255').Value }                        Else {" "}
                        'SubIFDs'                       = If ($image.Properties.Exists('330'))   { $image.Properties.Item('330').Value }                        Else {" "}
                        'SubjectDistance'               = If ($image.Properties.Exists('37382')) { $image.Properties.Item('37382').Value }                      Else {" "}
                        'SubjectLocation'               = If ($image.Properties.Exists('37396')) { $image.Properties.Item('37396').Value }                      Else {" "}
                        'SubjectLocation_2'             = If ($image.Properties.Exists('41492')) { $image.Properties.Item('41492').Value }                      Else {" "}
                        'SubSecTime'                    = If ($image.Properties.Exists('37520')) { $image.Properties.Item('37520').Value }                      Else {" "}
                        'SubSecTimeDigitized'           = If ($image.Properties.Exists('37522')) { $image.Properties.Item('37522').Value }                      Else {" "}
                        'SubSecTimeOriginal'            = If ($image.Properties.Exists('37521')) { $image.Properties.Item('37521').Value }                      Else {" "}
                        'SubTileBlockSize'              = If ($image.Properties.Exists('50974')) { $image.Properties.Item('50974').Value }                      Else {" "}
                        'T4Options'                     = If ($image.Properties.Exists('292'))   { $image.Properties.Item('292').Value }                        Else {" "}
                        'T6Options'                     = If ($image.Properties.Exists('293'))   { $image.Properties.Item('293').Value }                        Else {" "}
                        'TargetPrinter'                 = If ($image.Properties.Exists('337'))   { $image.Properties.Item('337').Value }                        Else {" "}
                        'Thresholding'                  = If ($image.Properties.Exists('263'))   { $image.Properties.Item('263').Value }                        Else {" "}
                        'TileByteCounts'                = If ($image.Properties.Exists('325'))   { $image.Properties.Item('325').Value }                        Else {" "}
                        'TileLength'                    = If ($image.Properties.Exists('323'))   { $image.Properties.Item('323').Value }                        Else {" "}
                        'TileOffsets'                   = If ($image.Properties.Exists('324'))   { $image.Properties.Item('324').Value }                        Else {" "}
                        'TileWidth'                     = If ($image.Properties.Exists('322'))   { $image.Properties.Item('322').Value }                        Else {" "}
                        'TimeZoneOffset'                = If ($image.Properties.Exists('34858')) { $image.Properties.Item('34858').Value }                      Else {" "}
                        'TransferFunction'              = If ($image.Properties.Exists('301'))   { $image.Properties.Item('301').Value }                        Else {" "}
                        'TransferRange'                 = If ($image.Properties.Exists('342'))   { $image.Properties.Item('342').Value }                        Else {" "}
                        'UniqueCameraModel'             = If ($image.Properties.Exists('50708')) { $image.Properties.Item('50708').Value }                      Else {" "}
                        'WhiteLevel'                    = If ($image.Properties.Exists('50717')) { $image.Properties.Item('50717').Value }                      Else {" "}
                        'WhitePoint'                    = If ($image.Properties.Exists('318'))   { (($image.Properties.Item('318').Value) | select -ExpandProperty Value) -join (" ") } Else {" "}
                        'XClipPathUnits'                = If ($image.Properties.Exists('344'))   { $image.Properties.Item('344').Value }                        Else {" "}
                        'YCbCrCoefficients'             = If ($image.Properties.Exists('529'))   { $image.Properties.Item('529').Value }                        Else {" "}
                        'YCbCrPositioning'              = If ($image.Properties.Exists('531'))   { $image.Properties.Item('531').Value }                        Else {" "}
                        'YCbCrSubSampling'              = If ($image.Properties.Exists('530'))   { $image.Properties.Item('530').Value }                        Else {" "}
                        'YClipPathUnits'                = If ($image.Properties.Exists('345'))   { $image.Properties.Item('345').Value }                        Else {" "}


                        # Credit: Franck Richard: "Use PowerShell to Remove Metadata and Resize Images": http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html
                        'Aperture'                      = If ($image.Properties.Exists('37378')) { ($image.Properties.Item('37378').Value).Value }                Else {" "}
                        'Color Space'                   = If ($image.Properties.Exists('40961')) { $image.Properties.Item('40961').Value }                      Else {" "}
                        'Contrast'                      = If ($image.Properties.Exists('41992')) { $contrast[[int]$image.Properties.Item('41992').Value] }      Else {" "}
                        'Custom Rendered'               = If ($image.Properties.Exists('41985')) { $custrender[[int]$image.Properties.Item('41985').Value] }    Else {" "}
                        'Date Created'                  = If ($image.Properties.Exists('36868')) { $image.Properties.Item('36868').Value }                      Else {" "}
                        'Date Taken'                    = If ($image.Properties.Exists('36867')) { $image.Properties.Item('36867').Value }                      Else {" "}
                        'Digital Zoom Ratio'            = If ($image.Properties.Exists('41988')) { ($image.Properties.Item('41988').Value).Value }              Else {" "}
                        'FocalLength'                   = If ($image.Properties.Exists('37386')) { ($image.Properties.Item('37386').Value).Value }              Else {" "}
                        'Equipment Maker'               = If ($image.Properties.Exists('271'))   { $image.Properties.Item('271').Value }                        Else {" "}
                        'Equipment Model'               = If ($image.Properties.Exists('272'))   { $image.Properties.Item('272').Value }                        Else {" "}
                        'Exposure Compensation'         = If ($image.Properties.Exists('37380')) { ($image.Properties.Item('37380').Value).Value }              Else {" "}
                        'Exposure Mode'                 = If ($image.Properties.Exists('41986')) { $exposure[[int]$image.Properties.Item('41986').Value] }      Else {" "}
                        'Exposure Time'                 = If ($image.Properties.Exists('33434')) { ($image.Properties.Item('33434').Value).Value }              Else {" "}
                        'F Number'                      = If ($image.Properties.Exists('33437')) { ($image.Properties.Item('33437').Value).Value }              Else {" "}
                        'File Source'                   = If ($image.Properties.Exists('41728')) { $filesrc[[int]$image.Properties.Item('41728').Value] }       Else {" "}
                    #    'Flash'                         = If ($image.Properties.Exists('37385')) { $image.Properties.Item('37385').Value }                      Else {" "}
                        'Flash'                         = If ($image.Properties.Exists('37385')) { $flash[[int]$image.Properties.Item('37385').Value] }         Else {" "}
                        'Focal Length in 35 mm Format'  = If ($image.Properties.Exists('41989')) { $image.Properties.Item('41989').Value }                      Else {" "}
                        'Focal Plane Resolution Unit'   = If ($image.Properties.Exists('41488')) { $focal[[int]$image.Properties.Item('41488').Value] }         Else {" "}
                        'Focal Plane X Resolution'      = If ($image.Properties.Exists('41486')) { $image.Properties.Item('41486').Value }                      Else {" "}
                        'Focal Plane Y Resolution'      = If ($image.Properties.Exists('41487')) { $image.Properties.Item('41487').Value }                      Else {" "}
                        'Gain Control'                  = If ($image.Properties.Exists('41991')) { $gain[[int]$image.Properties.Item('41991').Value] }          Else {" "}
                        'ISO Speed'                     = If ($image.Properties.Exists('34855')) { $image.Properties.Item('34855').Value }                      Else {" "}
                        'Maximum Aperture'              = If ($image.Properties.Exists('37381')) { ($image.Properties.Item('37381').Value).Value }              Else {" "}
                        'Metering Mode'                 = If ($image.Properties.Exists('37383')) { $metering[[int]$image.Properties.Item('37383').Value] }      Else {" "}
                        'Modified Date Time'            = If ($image.Properties.Exists('306'))   { $image.Properties.Item('306').Value }                        Else {" "}
                        'Image Orientation'             = If ($image.Properties.Exists('274'))   { $orient[[int]$image.Properties.Item('274').Value] }          Else {" "}
                        'Pixel X Dimension'             = If ($image.Properties.Exists('40962')) { $image.Properties.Item('40962').Value }                      Else {" "}
                        'Pixel Y Dimension'             = If ($image.Properties.Exists('40963')) { $image.Properties.Item('40963').Value }                      Else {" "}
                        'Resolution Unit'               = If ($image.Properties.Exists('296'))   { $unit[[int]$image.Properties.Item('296').Value] }            Else {" "}
                        'Saturation'                    = If ($image.Properties.Exists('41993')) { $saturation[[int]$image.Properties.Item('41993').Value] }    Else {" "}
                        'Scene Capture Type'            = If ($image.Properties.Exists('41990')) { $scene[[int]$image.Properties.Item('41990').Value] }         Else {" "}
                        'Sensing Method'                = If ($image.Properties.Exists('41495')) { $sensing[[int]$image.Properties.Item('41495').Value] }       Else {" "}
                        'Sharpness'                     = If ($image.Properties.Exists('41994')) { $sharpness[[int]$image.Properties.Item('41994').Value] }     Else {" "}
                        'Shutter Speed'                 = If ($image.Properties.Exists('37377')) { ($image.Properties.Item('37377').Value).Value }              Else {" "}
                        'Subject Distance Range'        = If ($image.Properties.Exists('41996')) { $sdr[[int]$image.Properties.Item('41996').Value] }           Else {" "}
                        'White Balance'                 = If ($image.Properties.Exists('41987')) { $white[[int]$image.Properties.Item('41987').Value] }         Else {" "}
                        'X Resolution'                  = If ($image.Properties.Exists('282'))   { ($image.Properties.Item('282').Value).Value }                Else {" "}
                        'Y Resolution'                  = If ($image.Properties.Exists('283'))   { ($image.Properties.Item('283').Value).Value }                Else {" "}

                } # New-Object
    } # ForEach ($file)




    # Step 3
    # Process the portrait pictures
    $new_portrait_picture_found = $false
    If (($vertical.Count -ge 1) -and (-not $ExcludePortrait)) {

                    # Increment the step counter
                    $task_number++

                    # Update the progress bar if there is more than one unique file to process
                    $activity = "Processing $($new_files.Count) Spotlight lock screen backgrounds - Step $task_number/$operations"
                    $task = "Processing the portrait pictures"
                    If (($new_files.Count) -ge 2) {
                        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($task_number / $total_steps) * 100)
                    } # If ($new_files.Count)

        # Copy the portrait pictures
        # Note: for WIA, please see above at Step 1 ("Load the image as an ImageFile object")
        # $image.SaveFile($destination_path)
        $portrait_candidates = $results | where { $_.Orientation -eq "Vertical" }
        ForEach ($picture in $portrait_candidates) {

            $new_portrait_picture_found = $true
            $image.LoadFile($picture.Source)
            $image.SaveFile($picture.Destination)
            $new_portrait += $picture

        } # ForEach $picture)
    } Else {
        $continue = $true
    } # Else (If $vertical.Count)




    # Step 4
    # Process the landscape wallpapers
    $new_landscape_wallpaper_found = $false
    If ($horizontal.Count -ge 1) {

                    # Increment the step counter
                    $task_number++

                    # Update the progress bar if there is more than one unique file to process
                    $activity = "Processing $($new_files.Count) Spotlight lock screen backgrounds - Step $task_number/$operations"
                    $task = "Processing the landscape wallpapers"
                    If (($new_files.Count) -ge 2) {
                        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($task_number / $total_steps) * 100)
                    } # If ($new_files.Count)

        # Copy the landscape wallpapers
        # Note: for WIA, please see above at Step 1 ("Load the image as an ImageFile object")
        $landscape_candidates = $results | where { $_.Orientation -eq "Horizontal" }
        ForEach ($wallpaper in $landscape_candidates) {

            $new_landscape_wallpaper_found = $true
            $image.LoadFile($wallpaper.Source)
            $image.SaveFile($wallpaper.Destination)
            $new_landscape += $wallpaper

        } # ForEach $wallpaper)
    } Else {
        $continue = $true
    } # Else (If $horizontal.Count)

                    # Close the progress bar if it has been opened
                    If (($new_files.Count) -gt 1) {
                        $task = "Finished retrieving Spotlight lock screen backgrounds."
                        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($total_steps / $total_steps) * 100) -Completed
                    } # If ($new_files.Count)

} # Process




End {

    <#
                    # Create an unique temp folder inside the destination folder with the NewGuid method    # Credit: robledosm: "Save lockscreen as wallpaper"
                    # Source: https://msdn.microsoft.com/en-us/library/system.guid.newguid(v=vs.110).aspx
                    # Source: https://blog.davidjwise.com/2012/02/22/creating-guids-in-powershell/
                    # $guid = [System.Guid]::NewGuid().Guid
                    # $temp = "$real_output_path\$guid"
                    # New-Item "$temp" -ItemType Directory -Force | Out-Null
    #>

    If ($results.Count -ge 1) {

        # Discard the headers that have empty values and display the new files in a pop-up window
        # Credit: Fred: "select-object | where": https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76ae6430-4993-4422-aa97-8f8ec3ca4e87/selectobject-where?forum=winserverpowershell
        $results.PSObject.TypeNames.Insert(0,"New Files")
        $results_selection = $results | Select-Object | ForEach-Object {
            $properties = $_.PsObject.Properties
            ForEach ($property in $properties) {
                If ($property.Value -eq " ") { $_.PSObject.Properties.Remove($property.Name) }
            } # ForEach
            $_
        } # ForEach-Object
        $results_selection | Sort Orientation,File | Out-GridView

                        # Display the results in console
                        $empty_line | Out-String
                        $results_console = $results | Sort Orientation,File | Select-Object 'File','Size','Type','Width','Height'
                        $results_console | Format-Table -Auto -Wrap

                        # Display rudimentary stats in console
                        If ($results.Count -ge 2) {
                            $text = "$($results.Count) new images found."
                        } ElseIf ($results.Count -eq 1) {
                            $text = "One new image found."
                        } Else {
                            $empty_line | Out-String
                            $text = "Didn't find any new images."
                        } # Else (If $results.Count)
                        Write-Output $text
                        $empty_line | Out-String

            # Sound the bell if set to do so with the -Audio parameter
            # Source: https://blogs.technet.microsoft.com/heyscriptingguy/2013/09/21/powertip-use-powershell-to-send-beep-to-console/
            If (($Audio) -and ($results.Count -ge 1)) {
                [console]::beep(2000,830)
            } Else {
                $continue = $true
            } # Else (If $Audio)

                        # Make the log entry if set to do so with the -Log parameter
                        # Note: Append parameter of Export-Csv was introduced in PowerShell 3.0.
                        # Source: http://stackoverflow.com/questions/21048650/how-can-i-append-files-using-export-csv-for-powershell-2
                        # Source: https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/
                        If (($Log) -and ($results.Count -ge 1)) {
                            $data = $results | Sort Orientation,File | Select-Object 'Source','File','Directory','BaseName','Original File Extension','FileExtension','Size','raw_size','Type','Orientation','Width','Height','Created','Modified','Saved to a New Location','Destination','HorizontalResolution','VerticalResolution','PixelDepth','FrameCount','ActiveFrame','IsIndexedPixelFormat','IsAlphaPixelFormat','IsExtendedPixelFormat','IsAnimated','Remarks','FormatID','SHA256','ActiveArea','AnalogBalance','AntiAliasStrength','Aperture','Artist','AsShotICCProfile','AsShotNeutral','AsShotPreProfileMatrix','AsShotWhiteXY','Author','BaselineExposure','BaselineNoise','BaselineSharpness','BatteryLevel','BayerGreenSplit','BestQualityScale','BitsPerSample','BlackLevel','BlackLevelDeltaH','BlackLevelDeltaV','BlackLevelRepeatDim','BodySerialNumber','BrightnessValue','Byte_AsShotProfileName','Byte_CameraCalibrationSignat','Byte_CFAPattern','Byte_CFAPlaneColor','Byte_ClipPath','Byte_DNGBackwardVersion','Byte_DNGPrivateData','Byte_DNGVersion','Byte_DotRange','Byte_GPSAltitudeRef','Byte_GPSVersionID','Byte_ImageResources','Byte_LocalizedCameraModel','Byte_OriginalRawFileName','Byte_PreviewApplicationName','Byte_PreviewApplicationVersion','Byte_PreviewSettingsDigest','Byte_PreviewSettingsName','Byte_ProfileCalibrationSignat','Byte_ProfileCopyright','Byte_ProfileName','Byte_RawDataUniqueID','Byte_TIFFEPStandardID','Byte_XMLPacket','CalibrationIlluminant1','CalibrationIlluminant2','CameraCalibration1','CameraCalibration2','CameraOwnerName','CameraSerialNumber','CellLength','CellWidth','CFALayout','CFAPattern','CFARepeatPatternDim','ChromaBlurRadius','Color Space','ColorimetricReference','ColorMap','ColorMatrix1','ColorMatrix2','Comment','ComponentsConfiguration','CompressedBitsPerPixel','Compression','Contrast','Copyright','CurrentICCProfile','CurrentPreProfileMatrix','Custom Rendered','Date Created','Date Taken','DefaultCropOrigin','DefaultCropSize','DefaultScale','DeviceSettingDescription','Digital Zoom Ratio','DocumentName','Equipment Maker','Equipment Model','ExifTag','ExifVersion','Exposure Compensation','Exposure Mode','Exposure Time','ExposureIndex','ExposureIndex_2','ExposureProgram','ExtraSamples','F Number','File Source','FillOrder','Flash','FlashEnergy','FlashEnergy_2','FlashpixVersion','Focal Length in 35 mm Format','Focal Plane Resolution Unit','Focal Plane X Resolution','Focal Plane Y Resolution','FocalLength','FocalPlaneResolutionUnit','FocalPlaneXResolution','FocalPlaneYResolution','ForwardMatrix1','ForwardMatrix2','Gain Control','GPSAltitude','GPSAreaInformation','GPSDateStamp','GPSDestBearing','GPSDestBearingRef','GPSDestDistance','GPSDestDistanceRef','GPSDestLatitude','GPSDestLatitudeRef','GPSDestLongitude','GPSDestLongitudeRef','GPSDifferential','GPSImgDirection','GPSImgDirectionRef','GPSLatitude','GPSLatitudeRef','GPSLongitude','GPSLongitudeRef','GPSMapDatum','GPSMeasureMode','GPSProcessingMethod','GPSSatellites','GPSSpeed','GPSSpeedRef','GPSStatus','GPSTag','GPSTimeStamp','GPSTrack','GPSTrackRef','GrayResponseCurve','GrayResponseUnit','HalftoneHints','HostComputer','Image Orientation','ImageDescription','ImageHistory','ImageID','ImageLength','ImageNumber','ImageUniqueID','ImageWidth','Indexed','InkNames','InkSet','Interlace','InteroperabilityTag','IPTCNAA','ISO Speed','ISOSpeed','ISOSpeedLatitudeyyy','ISOSpeedLatitudezzz','JPEGACTables','JPEGDCTables','JPEGInterchangeFormat','JPEGInterchangeFormatLength','JPEGLosslessPredictors','JPEGPointTransforms','JPEGProc','JPEGQTables','JPEGRestartInterval','JPEGTables','Keywords','LensInfo','LensMake','LensModel','LensSerialNumber','LensSpecification','LightSource','LinearizationTable','LinearResponseLimit','MakerNoteSafety','MaskedAreas','Maximum Aperture','Metering Mode','Modified Date Time','NewSubfileType','Noise','NoiseProfile','NoiseReductionApplied','NumberOfInks','OECF','OpcodeList1','OpcodeList2','OpcodeList3','OPIProxy','OriginalRawFileData','OriginalRawFileDigest','PageNumber','PhotometricInterpretation','Pixel X Dimension','Pixel Y Dimension','PlanarConfiguration','Predictor','PreviewColorSpace','PreviewDateTime','PrimaryChromaticities','PrintImageMatching','ProcessingSoftware','ProfileEmbedPolicy','ProfileHueSatMapData1','ProfileHueSatMapData2','ProfileHueSatMapDims','ProfileLookTableData','ProfileLookTableDims','ProfileToneCurve','Rating','RatingPercent','RawImageDigest','RecommendedExposureIndex','ReductionMatrix1','ReductionMatrix2','ReferenceBlackWhite','RelatedSoundFile','Resolution Unit','RowInterleaveFactor','RowsPerStrip','SampleFormat','SamplesPerPixel','Saturation','Scene Capture Type','SceneType','SecurityClassification','SelfTimerMode','Sensing Method','SensingMethod','SensitivityType','ShadowScale','Sharpness','Shutter Speed','SMaxSampleValue','SMinSampleValue','Software','SpatialFrequencyResponse_2','SpatialFrequencyResponse','SpectralSensitivity','StandardOutputSensitivity','StripByteCounts','StripOffsets','SubfileType','SubIFDs','Subject Distance Range','Subject','SubjectDistance','SubjectLocation_2','SubjectLocation','SubSecTime','SubSecTimeDigitized','SubSecTimeOriginal','SubTileBlockSize','T4Options','T6Options','TargetPrinter','Thresholding','TileByteCounts','TileLength','TileOffsets','TileWidth','TimeZoneOffset','Title','TransferFunction','TransferRange','UniqueCameraModel','User Comment','White Balance','WhiteLevel','WhitePoint','X Resolution','XClipPathUnits','Y Resolution','YCbCrCoefficients','YCbCrPositioning','YCbCrSubSampling','YClipPathUnits','File Original','File New'
                            $logfile_path = "$real_output_path\$log_filename"

                                If ((Test-Path $logfile_path) -eq $false) {
                                    $data | Export-Csv $logfile_path -Delimiter ';' -NoTypeInformation -Encoding UTF8
                                } Else {
                                    # $data | Export-Csv $logfile_path -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Append
                                    $data | ConvertTo-Csv -Delimiter ';' -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $logfile_path -Append -Encoding UTF8
                                    (Get-Content $logfile_path) | ForEach-Object { $_ -replace ('"', '') } | Out-File -FilePath $logfile_path -Force -Encoding UTF8
                                } # Else (If Test-Path $logfile_path)
                        } Else {
                            $continue = $true
                        } # Else (If $Log)

            # Open the -Output location in the File Manager, if set to do so with the -Open parameter
            If (($Open) -and ($Force)) {
                Invoke-Item $real_output_path
            } ElseIf (($Open) -and ($results.Count -ge 1)) {
                Invoke-Item $real_output_path
            } Else {
                $continue = $true
            } # Else (If $Open)

    } Else {
        $text = "Didn't process any Spotlight lock screen backgrounds. (Exit 3)"
        Write-Output $text
        $empty_line | Out-String
    } # Else (If $results.Count)

} # End




# [End of Line]


<#

   _____
  / ____|
 | (___   ___  _   _ _ __ ___ ___
  \___ \ / _ \| | | | '__/ __/ _ \
  ____) | (_) | |_| | | | (_|  __/
 |_____/ \___/ \__,_|_|  \___\___|


# Credit:               robledosm: "Save lockscreen as wallpaper"
# Forked from:          robledosm: "Save lockscreen as wallpaper"
# Original version:     robledosm: "Save lockscreen as wallpaper": https://github.com/robledosm/save-lockscreen-as-wallpaper

https://github.com/robledosm/save-lockscreen-as-wallpaper                                                   # robledosm: "Save lockscreen as wallpaper"
http://powershell.com/cs/media/p/7476.aspx                                                                  # clayman2: "Disk Space"
http://powershell.com/cs/forums/t/9685.aspx                                                                 # lamaar75: "Creating a Menu"
http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html                        # Franck Richard: "Use PowerShell to Remove Metadata and Resize Images"
http://stackoverflow.com/questions/14115415/call-windows-runtime-classes-from-powershell                    # klumsy: "Call Windows Runtime Classes from PowerShell"
https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76ae6430-4993-4422-aa97-8f8ec3ca4e87/selectobject-where?forum=winserverpowershell    # Fred: "select-object | where"
http://mikefrobbins.com/2015/03/31/powershell-advanced-functions-can-we-build-them-better-with-parameter-validation-yes-we-can/                     # Mike F Robbins: "PowerShell Advanced Functions: Can we build them better?"


  _    _      _
 | |  | |    | |
 | |__| | ___| |_ __
 |  __  |/ _ \ | '_ \
 | |  | |  __/ | |_) |
 |_|  |_|\___|_| .__/
               | |
               |_|
#>

<#
.SYNOPSIS
Retrieves Windows Spotlight lock screen wallpapers and saves them to a defined
directory.

.DESCRIPTION
 Get-Windows10LockScreenWallpapers uses by default one of three methods to determine
 the source path, where the Windows Spotlight lock screen wallpapers are stored
 locally:

 1. by reading the registry key
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Lock Screen\Creative\LandscapeAssetPath",
 2. by estimating the * value (and the source path) in
    "$($env:LOCALAPPDATA)\Packages\Microsoft.Windows.ContentDelivery*\LocalState\Assets"
    path, which on most Windows 10 machines would most likely point to the
    "\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets" directory, or
 3. by figuring out the current lock screen hive (which usually is in the
    $env:windir\Web\Screen directory).

The methods will be tested in an ascending order and selected as the primary (only)
method, if deemed to be valid. By adding the -Include parameter to the command
launching Get-Windows10LockScreenWallpapers the third method of wallpaper searching
will be enabled, so that Get-Windows10LockScreenWallpapers will also look to the
current lock screen hive, even if the first method (registry) or the second
(estimation) method was selected as the primary method for searching the available
local lock screen wallpapers.

Get-Windows10LockScreenWallpapers uses the inbuilt Get-FileHash cmdlet to calculate
SHA256 hash values of the files for determining, whether a wallpaper already exists
in the -Output folder or a portrait picture in the -Subfolder directory. By default
Get-Windows10LockScreenWallpapers writes the landscape files to
"$($env:USERPROFILE)\Pictures\Wallpapers"(, which is the default -Output directory),
and the portrait pictures are placed in a subfolder called "Vertical" inside the
folder specified with the -Output parameter. The save location ("destination") may
be set with the -Output parameter, and the name of the subfolder may be set with
the -Subfolder parameter - the former accepts a full path as a value, and the
latter just a plain directory name.

The images are loaded as ImageFile COM objects with Microsoft Windows Image
Acquisition (WIA, which relies on the Windows Image Acquisition (WIA) service
'stisvc'), and over 300 image properties (usually, though, most of them are empty
or non-existent...) are read from the pictures before the new images are copied to
their final destination, which is set with the -Output parameter. To exclude the
portrait pictures from the results altogether, the parameter -ExcludePortrait may
be added to the command launching Get-Windows10LockScreenWallpapers.

By using the -Force parameter the -Output directory will be created without asking
any questions, confirmations or additional selections (which will be prompted by
default, if the -Output path doesn't seem to exist), and if the -Open parameter
was used in adjunction with the -Force parameter in the command launching
Get-Windows10LockScreenWallpapers, the main destination path is opened in the
default File Manager in every case, regardless whether any new files were found
or not. The -Audio parameter will trigger an audible beep, if new files were
processed, and the -Log parameter will start a log creation procedure, in which
the extracted image properties are written to a new CSV file (spotlight_log.csv)
or appended to an existing log file. Please note that if any of the individual
parameter values include space characters, the individual value should be enclosed
in quotation marks (single or double), so that PowerShell can interpret the command
correctly. This script is forked from robledosm's script "Save lockscreen as wallpaper"
(https://github.com/robledosm/save-lockscreen-as-wallpaper).

.PARAMETER Output
with aliases -Path, -Destination and -OutputFolder. Specifies the primary folder,
where the acquired new images are to be saved, and defines the default location
to be used with the -Log parameter (spotlight_log.csv), and also sets the parent
directory for the -Subfolder parameter. The default save location for the horizontal
(landscape) wallpapers is "$($env:USERPROFILE)\Pictures\Wallpapers", which will be
used, if no value for the -Output parameter is defined in the command launching
Get-Windows10LockScreenWallpapers. For the best results in iterative usage of
Get-Windows10LockScreenWallpapers, the default value should remain constant and
be set according to the prevailing conditions (at line 18).

The value for the -Output parameter should be a valid file system path pointing
to a directory (a full path of a folder such as C:\Windows). In case the path
includes space characters, please enclose the path in quotation marks (single or
double). If the -Output parameter value seems to point to a non-existing directory,
the script will ask, whether the user wants to create the folder or not. This
query can be bypassed by using the -Force parameter. It's not mandatory to write
-Output in the get Windows 10 lock screen wallpapers command to invoke the -Output
parameter, as is shown in the Examples below.

.PARAMETER Subfolder
with aliases -SubfolderForThePortraitPictures, -SubfolderForTheVerticalPictures
and -SubfolderName. Specifies the name of the subfolder, where the new portrait
pictures are to be saved. If the -ExcludePortrait parameter is not used, the
subfolder directory will be located under the folder defined with the -Output
parameter. The value for the -Subfolder parameter should be a plain directory name
(omitting the path and the parent directories). The default value is "Vertical".
For the best results in iterative usage of Get-Windows10LockScreenWallpapers,
the default value should remain constant and be set according to the prevailing
conditions at line 30.

.PARAMETER Force
If the -Force parameter is added to the command launching
Get-Windows10LockScreenWallpapers, the -Output directory will be created without
asking any questions, confirmations or additional selections (which will be prompted
by default, if the -Output path doesn't seem to exist). If the -Open parameter is
used in adjunction with the -Force parameter in the command launching
Get-Windows10LockScreenWallpapers, the main destination path is opened in the
File Manager in every case, regardless whether any new files were found or not.

.PARAMETER ExcludePortrait
with aliases -NoPortrait, -NoSubfolder and -Exclude The -ExcludePortrait parameter
excludes all portrait (vertical) pictures from the files that will be copied to a
new location. Also prevents the (automatic) creation of the -Subfolder directory
inside the main output destination.

.PARAMETER Include
with an alias -IncludeCurrentLockScreenBackgroundHive. If the -Include parameter is
used in the command launching Get-Windows10LockScreenWallpapers, the third method
of wallpaper searching will be enabled, so that Get-Windows10LockScreenWallpapers
will also look to the current lock screen hive, even if the first method (registry)
or the second method (estimation) was selected as the primary method for searching
the available local lock screen wallpapers. Usually this will add a directory called
'$env:windir\Web\Screen' to the list of source paths to be queried for new images.
Please note that the items inside the current lock screen hive may be of varying
file extension type, including mostly .jpg and .png pictures.

.PARAMETER Log
If the -Log parameter is added to the command launching
Get-Windows10LockScreenWallpapers, a log file creation/updating procedure is
initiated, if new files are found. The log file (spotlight_log.csv) is created or
updated at the path defined with the -Output parameter. If the CSV log file seems
to already exist, new data will be appended to the end of that file. The log file
will consist over 300 image properties, of which most are empty.

The MakerNote Exif datafield is disabled (commented out in the source code) mainly
due to the vast variation of the formats used in the MakerNote Exif datafields
themselves. The exhaustive amount of different languages and data formats found in
the MakerNote Exif tags means that extensive additional coding efforts would be
needed for producing universally readable content from the MakerNote (37500) Exif
tag values.

The rather peculiar append procedure is used instead of the native -Append parameter
of the Export-Csv cmdlet for ensuring, that the CSV file will not contain any
additional quotation marks(, which might mess up the layout in some scenarios).

.PARAMETER Open
If the -Open parameter is used in the command launching
Get-Windows10LockScreenWallpapers and new files are found, the default File Manager
is opened at the -Output folder after the files are processed. If the -Force
parameter is used in adjunction with the -Open parameter, the main destination path
is opened in the File Manager in every case, regardless whether any new files were
found or not. Please note, though, that the -Force parameter will also Force the
creation of the -Output folder.

.PARAMETER Audio
If the -Audio parameter is used in the command launching
Get-Windows10LockScreenWallpapers and new files are found, an audible beep will occur.

.OUTPUTS
All new Windows Spotlight lock screen wallpapers are saved under a directory defined
with the -Output parameter. Displays wallpaper related info in console, and if any
new files were found, displays the results in a pop-up window (Out-GridView).
Optionally, if the -Log parameter was used in the command launching
Get-Windows10LockScreenWallpapers, and new files were found, a log file
(spotlight_log.csv) creation/updating procedure is initiated at the path defined
with the -Output variable. Also optionally, the default File Manager is opened at
the -Output folder, after the new files are processed, if the -Open parameter was
used in the command launching Get-Windows10LockScreenWallpapers. A progress bar is
also shown in console, if multiple images are being processed.


    Default values (the log file creation/updating procedure only occurs, if the
    -Log parameter is used and new files are found):

    $($env:USERPROFILE)\Pictures\Wallpapers                     : The folder for landscape wallpapers
    $($env:USERPROFILE)\Pictures\Wallpapers\spotlight_log.csv   : CSV-file
    $($env:USERPROFILE)\Pictures\Wallpapers\Vertical            : The folder for portrait pictures


.NOTES
Please note that all the parameters can be used in one get Windows 10 lock screen
wallpapers command and that each of the parameters can be "tab completed" before
typing them fully (by pressing the [tab] key).

For a non-commandline alternative, please see the SpotBright App at Microsoft Store.
https://www.microsoft.com/en-us/store/p/spotbright/9nblggh5km22

For instructions, how to enable the Windows Spotlight on the lock screen, please see
https://technet.microsoft.com/en-us/itpro/windows/manage/windows-spotlight or
http://www.windowscentral.com/how-enable-windows-spotlight

    Homepage:           https://github.com/auberginehill/get-windows-10-lock-screen-wallpapers
    Short URL:          http://tinyurl.com/z3ew947
    Version:            1.0

.EXAMPLE
./Get-Windows10LockScreenWallpapers
Runs the script. Please notice to insert ./ or .\ before the script name. If the
default -Output parameter folder "$($env:USERPROFILE)\Pictures\Wallpapers" doesn't
seem to exist, the user will be queried, whether the -Output folder should be
created or not. Will query "$($env:USERPROFILE)\Pictures\Wallpapers" and the default
subfolder ("$($env:USERPROFILE)\Pictures\Wallpapers\Vertical") for existing files,
and calculates their SHA256 hash values. Uses one of the three available methods
to retrieve the Windows Spotlight lock screen wallpapers and compares their SHA256
hash values against the hash values of the files in the -Output folder to determine,
whether any new files are present or not. All new landscape images are then copied
to the default -Output parameter folder ("$($env:USERPROFILE)\Pictures\Wallpapers"),
and all the new portrait images are copied to the default subfolder of the portrait
pictures ("$($env:USERPROFILE)\Pictures\Wallpapers\Vertical"). A pop-up window
showing the new files will open, if new files were found.

.EXAMPLE
help ./Get-Windows10LockScreenWallpapers -Full
Displays the help file.

.EXAMPLE
.\Get-Windows10LockScreenWallpapers.ps1 -Log -Audio -Open -ExcludePortrait -Force
Runs the script and creates the default -Output folder, if it doesn't exist, since
the -Force was used. Also, since the -Force was used, the File Manager will be
opened at the default -Output folder regardless whether any new files were found
or not. Uses one of the three available methods for retrieving the Windows Spotlight
lock screen wallpapers and compares their SHA256 hash values against the hash values
found in the default -Output folder to determine, whether any new files are present
or not. Since the -ExcludePortrait parameter was used, the results are limited to
the landscape wallpapers, and the vertical portrait pictures are excluded from the
images to be processed further. If new landscape (horizontal) images were found,
a log file creation/updating procedure is initiated, and a CSV-file
(spotlight_log.csv) is created/updated at the default -Output folder after the new
 landscape wallpapers are copied to their default destination folder. Furthermore,
 if new files were indeed found, an audible beep will occur.

.EXAMPLE
./Get-Windows10LockScreenWallpapers -Output C:\Users\Dropbox\ -Subfolder dc01 -Include
Uses one or two of the three available methods (registry, estimation and current
lock screen hive) as the basis for determining the source paths, where the Windows
Spotlight lock screen wallpapers are stored locally. Since the -Include parameter
was used, the third method of wallpaper searching will be used in any case, which
usually means that that the contents of '$env:windir\Web\Screen' are also used as
a source. Compares the SHA256 hash values of the found files against the hash values
found in the "C:\Users\Dropbox\" and "C:\Users\Dropbox\dc01" folders to determine,
whether any new files are present or not. All new landscape images are then copied
to the "C:\Users\Dropbox\" folder, and all the new portrait images are copied to
the "C:\Users\Dropbox\dc01" subfolder. Since the path or the subfolder name doesn't
contain any space characters, they don't needs to be enveloped with quotation marks.
Furthermore, the word -Output may be left out from the command as well, because
-Output values are read automatically from the first parameter position.

.EXAMPLE
Set-ExecutionPolicy remotesigned
This command is altering the Windows PowerShell rights to enable script execution
in the default (LocalMachine) scope, and defines the conditions under which Windows
PowerShell loads configuration files and runs scripts in general. In Windows Vista
and later versions of Windows, for running commands that change the execution policy
of the LocalMachine scope, Windows PowerShell has to be run with elevated rights
(run as an administrator). The default policy of the default (LocalMachine) scope is
"Restricted" and a command "Set-ExecutionPolicy Restricted", will "undo" the changes
made with the original example command above (had the policy not been changed before).
Execution policies for the local computer (LocalMachine) and for the current user
(CurrentUser) are stored in the registry (at for instance the
HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ExecutionPolicy key), and remain
effective until they are changed again. The execution policy for a particular session
(Process) is stored only in memory, and is discarded when the session is closed.


    Parameters:

    Restricted      Does not load configuration files or run scripts, but permits
                    individual commands. Restricted is the default execution policy.

    AllSigned       Scripts can run. Requires that all scripts and configuration
                    files be signed by a trusted publisher, including the scripts
                    that have been written on the local computer. Risks running
                    signed, but malicious, scripts.

    RemoteSigned    Requires a digital signature from a trusted publisher on scripts
                    and configuration files that are downloaded from the Internet
                    (including e-mail and instant messaging programs). Does not
                    require digital signatures on scripts that have been written on
                    the local computer. Permits running unsigned scripts that are
                    downloaded from the Internet, if the scripts are unblocked by
                    using the Unblock-File cmdlet. Risks running unsigned scripts
                    from sources other than the Internet and signed, but malicious,
                    scripts.

    Unrestricted    Loads all configuration files and runs all scripts.
                    Warns the user before running scripts and configuration files
                    that are downloaded from the Internet. Not only risks, but
                    actually permits, eventually, running any unsigned scripts from
                    any source. Risks running malicious scripts.

    Bypass          Nothing is blocked and there are no warnings or prompts.
                    This execution policy is designed for configurations, in
                    which a Windows PowerShell script is built in to a larger
                    application, or for configurations, in which Windows
                    PowerShell is the foundation for a program that has its own
                    security model. Not only risks, but actually permits running
                    any unsigned scripts from any source. Risks running malicious
                    scripts.

    Undefined       Removes the currently assigned execution policy from the current
                    scope. If the execution policy in all scopes is set to Undefined,
                    the effective execution policy is Restricted, which is the
                    default execution policy. This parameter will not alter or
                    remove the ("master") execution policy that is set with a Group
                    Policy setting.

    Please note, that the Group Policy setting "Turn on Script Execution" overrides
    the execution policies set in Windows PowerShell in all scopes. To find this
    ("master") setting, please, for example open the Group Policy Editor (gpedit.msc)
    and navigate to Computer Configuration > Administrative Templates >
    Windows Components > Windows PowerShell.


    Notes 	      - The Group Policy setting ("Turn on Script Execution") is present
                    in the Windows Server 2003 Service Pack 1 or later, and on the
                    consumer products (depending on the Windows edition) in
                    Windows XP Service Pack 2 or later.

                  - The Group Policy Editor is not available in any Home or Starter
                    editions of Windows, be it Windows XP, Windows 7, Windows 8.1
                    or Windows 10.

                  - Group Policy (gpedit.msc) setting "Turn on Script Execution":
                    Not configured 	                                          : No effect, the default
                                                                                value of this setting
                    Disabled 	                                              : Restricted
                    Enabled - Allow only signed scripts 	                  : AllSigned
                    Enabled - Allow local scripts and remote signed scripts   : RemoteSigned
                    Enabled - Allow all scripts 	                          : Unrestricted


For more information, please type "Get-ExecutionPolicy -List", "help Set-ExecutionPolicy -Full",
"help about_Execution_Policies" or visit https://technet.microsoft.com/en-us/library/hh849812.aspx
or http://go.microsoft.com/fwlink/?LinkID=135170.

.EXAMPLE
New-Item -ItemType File -Path C:\Temp\Get-Windows10LockScreenWallpapers.ps1
Creates an empty ps1-file to the C:\Temp directory. The New-Item cmdlet has an inherent -NoClobber mode
built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing
file is about to happen. Overwriting a file with the New-Item cmdlet requires using the Force. If the
path name and/or the filename includes space characters, please enclose the whole -Path parameter value
in quotation marks (single or double):

    New-Item -ItemType File -Path "C:\Folder Name\Get-Windows10LockScreenWallpapers.ps1"

For more information, please type "help New-Item -Full".

.LINK
https://github.com/robledosm/save-lockscreen-as-wallpaper
http://powershell.com/cs/media/p/7476.aspx
http://powershell.com/cs/forums/t/9685.aspx
http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html
http://stackoverflow.com/questions/14115415/call-windows-runtime-classes-from-powershell
https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76ae6430-4993-4422-aa97-8f8ec3ca4e87/selectobject-where?forum=winserverpowershell
http://mikefrobbins.com/2015/03/31/powershell-advanced-functions-can-we-build-them-better-with-parameter-validation-yes-we-can/
http://www.askvg.com/windows-10-wallpapers-and-lock-screen-backgrounds/
https://www.cnet.com/how-to/where-to-find-the-windows-spotlight-photos/
http://www.windowscentral.com/how-save-all-windows-spotlight-lockscreen-images
https://github.com/hashhar/Windows-Hacks/blob/master/scheduled-tasks/save-windows-spotlight-lockscreens/save-spotlight.ps1
https://answers.microsoft.com/en-us/insider/forum/insider_wintp-insider_personal/windows-10041-windows-spotlight-lock-screen/5b1cddaf-7057-443b-99b6-8c3486a75262
http://www.winhelponline.com/blog/find-file-name-lock-screen-image-current-displayed/
https://www.tenforums.com/tutorials/38717-windows-spotlight-background-images-find-save-windows-10-a.html
http://www.ohmancorp.com/RefWin-windows8-change-pre-login-screen-background.asp
https://social.technet.microsoft.com/Forums/windows/en-US/a8db890c-204f-404a-bf74-3aa4c895b183/cant-customize-lock-or-logon-screen-backgrounds?forum=W8ITProPreRel
https://technet.microsoft.com/en-us/library/ff730939.aspx
https://technet.microsoft.com/en-us/library/ee692804.aspx
http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string#32608908
https://technet.microsoft.com/en-us/library/ee692803.aspx
https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.utility/get-filehash
http://nicholasarmstrong.com/2010/02/exif-quick-reference/
https://msdn.microsoft.com/en-us/library/ms630826(v=vs.85).aspx
http://kb.winzip.com/kb/entry/207/
https://msdn.microsoft.com/en-us/library/windows/desktop/ms630506(v=vs.85).aspx
https://blogs.msdn.microsoft.com/powershell/2009/03/30/image-manipulation-in-powershell/
http://stackoverflow.com/questions/4304821/get-startup-type-of-windows-service-using-powershell
https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.management/get-wmiobject
https://social.microsoft.com/Forums/en-US/4dfe4eec-2b9b-4e6e-a49e-96f5a108c1c8/using-powershell-as-a-photoshop-replacement?forum=Offtopic
https://msdn.microsoft.com/en-us/library/system.io.path_methods(v=vs.110).aspx
https://msdn.microsoft.com/en-us/library/ms630826(VS.85).aspx#SharedSample012
https://www.experts-exchange.com/questions/25100459/I-need-to-send-the-details-of-a-jpg-file-to-an-array-any-windows-api-to-do-this-or-get-me-started.html
https://social.technet.microsoft.com/Forums/windowsserver/en-US/16124c53-4c7f-41f2-9a56-7808198e102a/attribute-seems-to-give-byte-array-how-to-convert-to-string?forum=winserverpowershell
http://compgroups.net/comp.databases.ms-access/handy-routine-for-getting-file-metad/1484921
http://www.exiv2.org/tags.html
https://blogs.technet.microsoft.com/heyscriptingguy/2013/09/21/powertip-use-powershell-to-send-beep-to-console/
http://stackoverflow.com/questions/21048650/how-can-i-append-files-using-export-csv-for-powershell-2
https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/

#>

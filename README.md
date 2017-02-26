<!-- Visual Studio Code: For a more comfortable reading experience, use the key combination Ctrl + Shift + V
     Visual Studio Code: To crop the tailing end space characters out, please use the key combination Ctrl + A Ctrl + K Ctrl + X (Formerly Ctrl + Shift + X)
     Visual Studio Code: To improve the formatting of HTML code, press Shift + Alt + F and the selected area will be reformatted in a html file.
     Visual Studio Code shortcuts: http://code.visualstudio.com/docs/customization/keybindings (or https://aka.ms/vscodekeybindings)
     Visual Studio Code shortcut PDF (Windows): https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf




   _____      _       __          ___           _                  __  ___  _                _     _____                      __          __   _ _
  / ____|    | |      \ \        / (_)         | |                /_ |/ _ \| |              | |   / ____|                     \ \        / /  | | |
 | |  __  ___| |_ _____\ \  /\  / / _ _ __   __| | _____      _____| | | | | |     ___   ___| | _| (___   ___ _ __ ___  ___ _ _\ \  /\  / /_ _| | |_ __   __ _ _ __   ___ _ __ ___
 | | |_ |/ _ \ __|______\ \/  \/ / | | '_ \ / _` |/ _ \ \ /\ / / __| | | | | |    / _ \ / __| |/ /\___ \ / __| '__/ _ \/ _ \ '_ \ \/  \/ / _` | | | '_ \ / _` | '_ \ / _ \ '__/ __|
 | |__| |  __/ |_        \  /\  /  | | | | | (_| | (_) \ V  V /\__ \ | |_| | |___| (_) | (__|   < ____) | (__| | |  __/  __/ | | \  /\  / (_| | | | |_) | (_| | |_) |  __/ |  \__ \
  \_____|\___|\__|        \/  \/   |_|_| |_|\__,_|\___/ \_/\_/ |___/_|\___/|______\___/ \___|_|\_\_____/ \___|_|  \___|\___|_| |_|\/  \/ \__,_|_|_| .__/ \__,_| .__/ \___|_|  |___/
                                                                                                                                                  | |         | |
                                                                                                                                                  |_|         |_|                               -->


## Get-Windows10LockScreenWallpapers.ps1


<table>
    <tr>
        <td style="padding:6px"><strong>OS:</strong></td>
        <td colspan="2" style="padding:6px">Windows</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Type:</strong></td>
        <td colspan="2" style="padding:6px">A Windows PowerShell script</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Language:</strong></td>
        <td colspan="2" style="padding:6px">Windows PowerShell</td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Description:</strong></td>
        <td colspan="2" style="padding:6px">
            <p>
                Get-Windows10LockScreenWallpapers uses by default one of the three methods listed below to determine the source path, where the Windows Spotlight lock screen wallpapers are stored locally:</p>
                <ol>
                    <li>Reading the registry key "<code>HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Lock&nbsp;Screen\Creative\LandscapeAssetPath</code>"</li>
                    <li>Estimating the * value (and the source path) in "<code>$($env:LOCALAPPDATA)\Packages\Microsoft.Windows.ContentDelivery*\LocalState\Assets</code>" path, which on most Windows 10 machines would most likely point to the "<code>\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets</code>" directory</li>
                    <li>Figuring out the current lock screen hive (which usually is in the <code>$env:windir\Web\Screen</code> directory)</li>
                </ol>
            <p>                
                The methods are tested in an ascending order and selected as the primary (only) method, if deemed to be valid. By adding the <code>-Include</code> parameter to the command launching Get-Windows10LockScreenWallpapers the third method of wallpaper searching will be enabled, so that Get-Windows10LockScreenWallpapers will also look to the current lock screen hive, even if the first method (registry) or the second (estimation) method was selected as the primary method for searching the available local lock screen wallpapers.</p>
            <p>
                Get-Windows10LockScreenWallpapers uses the inbuilt <code>Get-FileHash</code> cmdlet to calculate SHA256 hash values of the files for determining, whether a wallpaper already exists in the <code>-Output</code> folder or a portrait picture in the <code>-Subfolder</code> directory. By default Get-Windows10LockScreenWallpapers writes the landscape files to "<code>$($env:USERPROFILE)\Pictures\Wallpapers</code>"(, which is the default <code>-Output</code> directory), and the portrait pictures are placed in a subfolder called "<code>Vertical</code>" inside the folder specified with the <code>-Output</code> parameter. The primary save location ("<dfn>destination</dfn>") may be set with the <code>-Output</code> parameter, and the name of the subfolder may be set with the <code>-Subfolder</code> parameter – the former accepts a full path as a value, and the latter just a plain directory name.</p>
            <p>
                The images are loaded as ImageFile COM objects with Microsoft Windows Image Acquisition (WIA, which relies on the Windows Image Acquisition (WIA) service '<code>stisvc</code>'), and over 300 image properties (usually, though, most of them are empty or non-existent...) are read from the pictures before the new images are copied to their final destination. To exclude the portrait pictures from the results altogether, the parameter <code>-ExcludePortrait</code> may be added to the command launching Get-Windows10LockScreenWallpapers.</p>
            <p>
                By using the <code>-Force</code> parameter the <code>-Output</code> directory will be created without asking any questions, confirmations or additional selections (which will be prompted by default, if the <code>-Output</code> path doesn't seem to exist), and if the <code>-Open</code> parameter was used in adjunction with the <code>-Force</code> parameter in the command launching Get-Windows10LockScreenWallpapers, the main destination path is opened in the default File Manager in every case, regardless whether any new files were found or not. The <code>-Audio</code> parameter will trigger an audible beep, if new files were processed, and the <code>-Log</code> parameter will start a log creation procedure, in which the extracted image properties are written to a new CSV file (<code>spotlight_log.csv</code>) or appended to an existing log file. Please note that if any of the individual parameter values include space characters, the individual value should be enclosed in quotation marks (single or double), so that PowerShell can interpret the command correctly. This script is forked from robledosm's script <a href="https://github.com/robledosm/save-lockscreen-as-wallpaper">Save lockscreen as wallpaper</a>.</p>
        </td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Homepage:</strong></td>
        <td colspan="2" style="padding:6px"><a href="https://github.com/auberginehill/get-windows-10-lock-screen-wallpapers">https://github.com/auberginehill/get-windows-10-lock-screen-wallpapers</a>
            <br />Short URL: <a href="http://tinyurl.com/z3ew947">http://tinyurl.com/z3ew947</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Version:</strong></td>
        <td colspan="2" style="padding:6px">1.0</td>
    </tr>
    <tr>
        <td rowspan="8" style="padding:6px"><strong>Sources:</strong></td>
        <td style="padding:6px">Emojis:</td>
        <td style="padding:6px"><a href="https://github.com/auberginehill/emoji-table">Emoji Table</a></td>
    </tr>
    <tr>
        <td style="padding:6px">robledosm:</td>
        <td style="padding:6px"><a href="https://github.com/robledosm/save-lockscreen-as-wallpaper">Save lockscreen as wallpaper</a></td>
    </tr>
    <tr>
        <td style="padding:6px">clayman2:</td>
        <td style="padding:6px"><a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a> (or one of the <a href="http://web.archive.org/web/20120304222258/http://powershell.com/cs/media/p/7476.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">lamaar75:</td>
        <td style="padding:6px"><a href="http://powershell.com/cs/forums/t/9685.aspx">Creating a Menu</a> (or one of the <a href="https://web.archive.org/web/20150910111758/http://powershell.com/cs/forums/t/9685.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">Franck Richard:</td>
        <td style="padding:6px"><a href="http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html">Use PowerShell to Remove Metadata and Resize Images</a></td>
    </tr>
    <tr>
        <td style="padding:6px">klumsy:</td>
        <td style="padding:6px"><a href="http://stackoverflow.com/questions/14115415/call-windows-runtime-classes-from-powershell">Call Windows Runtime Classes from PowerShell</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Fred:</td>
        <td style="padding:6px"><a href="https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76ae6430-4993-4422-aa97-8f8ec3ca4e87/selectobject-where?forum=winserverpowershell">select-object | where</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Mike F Robbins:</td>
        <td style="padding:6px"><a href="http://mikefrobbins.com/2015/03/31/powershell-advanced-functions-can-we-build-them-better-with-parameter-validation-yes-we-can/">PowerShell Advanced Functions: Can we build them better?</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><strong>Downloads:</strong></td>
        <td colspan="2" style="padding:6px">For instance <a href="https://raw.githubusercontent.com/auberginehill/get-windows-10-lock-screen-wallpapers/master/Get-Windows10LockScreenWallpapers.ps1">Get-Windows10LockScreenWallpapers.ps1</a>.
            Or <a href="https://github.com/auberginehill/get-windows-10-lock-screen-wallpapers/archive/master.zip">everything as a .zip-file</a>.</td>
    </tr>
</table>




### Screenshot

<ul><ul><ul>
<img class="screenshot" title="screenshot" alt="screenshot" height="80%" width="80%" src="https://raw.githubusercontent.com/auberginehill/get-windows-10-lock-screen-wallpapers/master/Get-Windows10LockScreenWallpapers.png">
</ul></ul></ul>




### Parameters

<table>
    <tr>
        <th>:triangular_ruler:</th>
        <td style="padding:6px">
            <ul>
                <li>
                    <h5>Parameter <code>-Output</code></h5>
                    <p>with aliases <code>-Path</code>, <code>-Destination</code> and <code>-OutputFolder</code>. Specifies the primary folder, where the acquired new images are to be saved, and defines the default location to be used with the <code>-Log</code> parameter (<code>spotlight_log.csv</code>), and also sets the parent directory for the <code>-Subfolder</code> parameter. The default save location for the horizontal (landscape) wallpapers is "<code>$($env:USERPROFILE)\Pictures\Wallpapers</code>", which will be used, if no value for the <code>-Output</code> parameter is defined in the command launching Get-Windows10LockScreenWallpapers. For the best results in iterative usage of Get-Windows10LockScreenWallpapers, the default value should remain constant and be set according to the prevailing conditions (at line 18).</p>
                    <p>The value for the <code>-Output</code> parameter should be a valid file system path pointing to a directory (a full path of a folder such as <code>C:\Windows</code>). In case the path includes space characters, please enclose the path in quotation marks (single or double). If the <code>-Output</code> parameter value seems to point to a non-existing directory, the script will ask, whether the user wants to create the folder or not. This query can be bypassed by using the <code>-Force</code> parameter. It's not mandatory to write <code>-Output</code> in the get Windows 10 lock screen wallpapers command to invoke the <code>-Output</code> parameter, as is described in the Examples below.</p>
                </li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>
                        <h5>Parameter <code>-Subfolder</code></h5>
                        <p>with aliases <code>-SubfolderForThePortraitPictures</code>, <code>-SubfolderForTheVerticalPictures</code> and <code>-SubfolderName</code>. Specifies the name of the subfolder, where the new portrait pictures are to be saved. If the <code>-ExcludePortrait</code> parameter is not used, the subfolder directory will be located under the folder defined with the <code>-Output</code> parameter. The value for the <code>-Subfolder</code> parameter should be a plain directory name (omitting the path and the parent directories). The default value is "<code>Vertical</code>". For the best results in iterative usage of Get-Windows10LockScreenWallpapers, the default value should remain constant and be set according to the prevailing conditions at line 30. </p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Force</code></h5>
                        <p>If the <code>-Force</code> parameter is added to the command launching Get-Windows10LockScreenWallpapers, the <code>-Output</code> directory will be created without asking any questions, confirmations or additional selections (which will be prompted by default, if the <code>-Output</code> path doesn't seem to exist). If the <code>-Open</code> parameter is used in adjunction with the <code>-Force</code> parameter in the command launching Get-Windows10LockScreenWallpapers, the main destination path is opened in the File Manager in every case, regardless whether any new files were found or not.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-ExcludePortrait</code></h5>
                        <p>with aliases <code>-NoPortrait</code>, <code>-NoSubfolder</code> and <code>-Exclude</code> The <code>-ExcludePortrait</code> parameter excludes all portrait (vertical) pictures from the files that will be copied to a new location. <code>-ExcludePortrait</code> also prevents the (automatic) creation of the <code>-Subfolder</code> directory inside the main output destination.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Include</code></h5>
                        <p>with an alias <code>-IncludeCurrentLockScreenBackgroundHive</code>. If the <code>-Include</code> parameter is used in the command launching Get-Windows10LockScreenWallpapers, the third method of wallpaper searching will be enabled, so that Get-Windows10LockScreenWallpapers will also look to the current lock screen hive, even if the first method (registry) or the second method (estimation) was selected as the primary method for searching the available local lock screen wallpapers. Usually this will add a directory called '<code>$env:windir\Web\Screen</code>' to the list of source paths to be queried for new images. Please note that the items inside the current lock screen hive may be of varying file extension type, including mostly <code>.jpg</code> and <code>.png</code> pictures.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Log</code></h5>
                        <p>If the <code>-Log</code> parameter is added to the command launching Get-Windows10LockScreenWallpapers, a log file creation/updating procedure is initiated, if new files are found. The log file (<code>spotlight_log.csv</code>) is created or updated at the path defined with the <code>-Output</code> parameter. If the CSV log file seems to already exist, new data will be appended to the end of that file. The log file will consist over 300 image properties, of which most are empty.</p>
                        <p>The <code>MakerNote</code> Exif datafield is disabled (commented out in the source code) mainly due to the vast variation of the formats used in the <code>MakerNote</code> Exif datafields themselves. The exhaustive amount of different languages and data formats found in the <code>MakerNote</code> Exif tags means that extensive additional coding efforts would be needed for producing universally readable content from the <code>MakerNote</code> (37500) Exif tag values.</p>
                        <p>A rather peculiar append procedure is used instead of the native <code>-Append</code> parameter of the <code>Export-Csv</code> cmdlet for ensuring, that the CSV file will not contain any additional quotation marks(, which might mess up the layout in some scenarios).</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Open</code></h5>
                        <p>If the <code>-Open</code> parameter is used in the command launching Get-Windows10LockScreenWallpapers and new files are found, the default File Manager is opened at the <code>-Output</code> folder after the files are processed. If the <code>-Force</code> parameter is used in adjunction with the <code>-Open</code> parameter, the main destination path is opened in the File Manager regardless whether any new files were found or not. Please note, though, that the <code>-Force</code> parameter will also Force the creation of the <code>-Output</code> folder.</p>
                    </li>
                </p>
                <p>
                    <li>
                        <h5>Parameter <code>-Audio</code></h5>
                        <p>If the <code>-Audio</code> parameter is used in the command launching Get-Windows10LockScreenWallpapers and new files are found, an audible beep will occur.</p>
                    </li>
                </p>
            </ul>
        </td>
    </tr>
</table>




### Outputs

<table>
    <tr>
        <th>:arrow_right:</th>
        <td style="padding:6px">
            <ul>
                <li>All new Windows Spotlight lock screen wallpapers are saved under a directory defined with the <code>-Output</code> parameter. Displays wallpaper related info in console, and if any new files were found, displays the results in a pop-up window (<code>Out-GridView</code>). Optionally, if the <code>-Log</code> parameter was used in the command launching Get-Windows10LockScreenWallpapers, and new files were found, a log file (<code>spotlight_log.csv</code>) creation/updating procedure is initiated at the path defined with the <code>-Output</code> variable. Also optionally, the default File Manager is opened at the <code>-Output</code> folder, after the new files are processed, if the <code>-Open</code> parameter was used in the command launching Get-Windows10LockScreenWallpapers. A progress bar is also shown in console, if multiple images are being processed.</li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>Default values (the log file creation/updating procedure only occurs, if the <code>-Log</code> parameter is used and new files are found):</li>
                </p>
                <ol>
                    <p>
                        <table>
                            <tr>
                                <td style="padding:6px"><strong>Path</strong></td>
                                <td style="padding:6px"><strong>Type</strong></td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$($env:USERPROFILE)\Pictures\Wallpapers</code></td>
                                <td style="padding:6px">The folder for landscape wallpapers</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$($env:USERPROFILE)\Pictures\Wallpapers\spotlight_log.csv</code></td>
                                <td style="padding:6px">CSV-file</td>
                            </tr>
                            <tr>
                                <td style="padding:6px"><code>$($env:USERPROFILE)\Pictures\Wallpapers\Vertical</code></td>
                                <td style="padding:6px">The folder for portrait pictures</td>
                            </tr>
                        </table>
                    </p>
                </ol>
            </ul>
        </td>
    </tr>
</table>




### Notes

<table>
    <tr>
        <th>:warning:</th>
        <td style="padding:6px">
            <ul>
                <li>Please note that all the parameters can be used in one get Windows 10 lock screen wallpapers command and that each of the parameters can be "tab completed" before typing them fully (by pressing the <code>[tab]</code> key).</li>
            </ul>
        </td>
    </tr>
    <tr>
        <th></th>
        <td style="padding:6px">
            <ul>
                <p>
                    <li>For a non-commandline alternative, please see the <a href="https://www.microsoft.com/en-us/store/p/spotbright/9nblggh5km22">SpotBright App</a> at Microsoft Store.</li>
                    <li>For instructions, how to enable the Windows Spotlight on the lock screen, please see <a href="https://technet.microsoft.com/en-us/itpro/windows/manage/windows-spotlight">Windows Spotlight on the lock screen</a> or <a href="http://www.windowscentral.com/how-enable-windows-spotlight">How to enable Windows spotlight in Windows 10</a>.</li>
                </p>
            </ul>
        </td>
    </tr>
</table>




### Examples

<table>
    <tr>
        <th>:book:</th>
        <td style="padding:6px">To open this code in Windows PowerShell, for instance:</td>
   </tr>
   <tr>
        <th></th>
        <td style="padding:6px">
            <ol>
                <p>
                    <li><code>./Get-Windows10LockScreenWallpapers</code><br />
                    Runs the script. Please notice to insert <code>./</code> or <code>.\</code> before the script name. If the default <code>-Output</code> parameter folder "<code>$($env:USERPROFILE)\Pictures\Wallpapers</code>" doesn't seem to exist, the user will be queried, whether the <code>-Output</code> folder should be created or not. Will query "<code>$($env:USERPROFILE)\Pictures\Wallpapers</code>" and the default subfolder ("<code>$($env:USERPROFILE)\Pictures\Wallpapers\Vertical</code>") for existing files, and calculates their SHA256 hash values. Uses one of the three available methods to retrieve the Windows Spotlight lock screen wallpapers and compares their SHA256 hash values against the hash values of the files in the <code>-Output</code> folder to determine, whether any new files are present or not. All new landscape images are then copied to the default <code>-Output</code> parameter folder ("<code>$($env:USERPROFILE)\Pictures\Wallpapers</code>"), and all the new portrait images are copied to the default subfolder of the portrait pictures ("<code>$($env:USERPROFILE)\Pictures\Wallpapers\Vertical</code>"). A pop-up window showing the new files will open, if new files were found.</li>
                </p>
                <p>
                    <li><code>help ./Get-Windows10LockScreenWallpapers -Full</code><br />
                    Displays the help file.</li>
                </p>
                <p>
                    <li><code>./Get-Windows10LockScreenWallpapers.ps1 -Log -Audio -Open -ExcludePortrait -Force</code><br />
                    Runs the script and creates the default <code>-Output</code> folder, if it doesn't exist, since the <code>-Force</code> was used. Also, since the <code>-Force</code> was used, the File Manager will be opened at the default <code>-Output</code> folder regardless whether any new files were found or not. Uses one of the three available methods for retrieving the Windows Spotlight lock screen wallpapers and compares their SHA256 hash values against the hash values found in the default <code>-Output</code> folder to determine, whether any new files are present or not. Since the <code>-ExcludePortrait</code> parameter was used, the results are limited to the landscape wallpapers, and the vertical portrait pictures are excluded from the images to be processed further. If new landscape (horizontal) images were found, after the new landscape wallpapers are copied to their default destination, a log file creation/updating procedure is initiated, and a CSV-file (<code>spotlight_log.csv</code>) is created/updated at the default <code>-Output</code> folder. Furthermore, if new files were indeed found, an audible beep will occur.</li>
                </p>
                <p>
                    <li><code>./Get-Windows10LockScreenWallpapers <code>-Output</code> C:\Users\Dropbox\ <code>-Subfolder</code> dc01 -Include</code><br />
                    Uses one or two of the three available methods (registry, estimation and current lock screen hive) as the basis for determining the source paths, where the Windows Spotlight lock screen wallpapers are stored locally. Since the <code>-Include</code> parameter was used, the third method of wallpaper searching will be used in any case, which usually means that that the contents of '<code>$env:windir\Web\Screen</code>' are also used as a source. Compares the SHA256 hash values of the found files against the hash values found in the "<code>C:\Users\Dropbox\</code>" and "<code>C:\Users\Dropbox\dc01</code>" folders to determine, whether any new files are present or not. All new landscape images are then copied to the "<code>C:\Users\Dropbox\</code>" folder, and all the new portrait images are copied to the "<code>C:\Users\Dropbox\dc01</code>" subfolder. Since the path or the subfolder name doesn't contain any space characters, they don't needs to be enveloped with quotation marks. Furthermore, the word <code>-Output</code> may be left out from the command as well, because <code>-Output</code> values are read automatically from the first parameter position.</li>
                </p>
                <p>
                    <li><p><code>Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine</code><br />
                    This command is altering the Windows PowerShell rights to enable script execution in the default (<code>LocalMachine</code>) scope, and defines the conditions under which Windows PowerShell loads configuration files and runs scripts in general. In Windows Vista and later versions of Windows, for running commands that change the execution policy of the <code>LocalMachine</code> scope, Windows PowerShell has to be run with elevated rights (<dfn>Run as Administrator</dfn>). The default policy of the default (<code>LocalMachine</code>) scope is "<code>Restricted</code>", and a command "<code>Set-ExecutionPolicy Restricted</code>" will "<dfn>undo</dfn>" the changes made with the original example above (had the policy not been changed before). Execution policies for the local computer (<code>LocalMachine</code>) and for the current user (<code>CurrentUser</code>) are stored in the registry (at for instance the <code>HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ExecutionPolicy</code> key), and remain effective until they are changed again. The execution policy for a particular session (<code>Process</code>) is stored only in memory, and is discarded when the session is closed.</p>
                        <p>Parameters:
                            <ul>
                                <table>
                                    <tr>
                                        <td style="padding:6px"><code>Restricted</code></td>
                                        <td colspan="2" style="padding:6px">Does not load configuration files or run scripts, but permits individual commands. <code>Restricted</code> is the default execution policy.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>AllSigned</code></td>
                                        <td colspan="2" style="padding:6px">Scripts can run. Requires that all scripts and configuration files be signed by a trusted publisher, including the scripts that have been written on the local computer. Risks running signed, but malicious, scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>RemoteSigned</code></td>
                                        <td colspan="2" style="padding:6px">Requires a digital signature from a trusted publisher on scripts and configuration files that are downloaded from the Internet (including e-mail and instant messaging programs). Does not require digital signatures on scripts that have been written on the local computer. Permits running unsigned scripts that are downloaded from the Internet, if the scripts are unblocked by using the <code>Unblock-File</code> cmdlet. Risks running unsigned scripts from sources other than the Internet and signed, but malicious, scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>Unrestricted</code></td>
                                        <td colspan="2" style="padding:6px">Loads all configuration files and runs all scripts. Warns the user before running scripts and configuration files that are downloaded from the Internet. Not only risks, but actually permits, eventually, running any unsigned scripts from any source. Risks running malicious scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>Bypass</code></td>
                                        <td colspan="2" style="padding:6px">Nothing is blocked and there are no warnings or prompts. Not only risks, but actually permits running any unsigned scripts from any source. Risks running malicious scripts.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px"><code>Undefined</code></td>
                                        <td colspan="2" style="padding:6px">Removes the currently assigned execution policy from the current scope. If the execution policy in all scopes is set to <code>Undefined</code>, the effective execution policy is <code>Restricted</code>, which is the default execution policy. This parameter will not alter or remove the ("<dfn>master</dfn>") execution policy that is set with a Group Policy setting.</td>
                                    </tr>
                                    <tr>
                                        <td colspan="3" style="padding:6px">Please note, that the Group Policy setting "<code>Turn on Script Execution</code>" overrides the execution policies set in Windows PowerShell in all scopes. To find this ("<dfn>master</dfn>") setting, please, for example, open the Group Policy Editor (<code>gpedit.msc</code>) and navigate to Computer Configuration → Administrative Templates → Windows Components → Windows PowerShell.</td>
                                    </tr>
                                    <tr>
                                        <td style="padding:6px">Notes</td>
                                        <td colspan="2" style="padding:6px">
                                            <ul>
                                                <li>The Group Policy Editor is not available in any Home or Starter editions of Windows, be it Windows XP, Windows 7, Windows 8.1 or Windows 10.</li>
                                            </ul>
                                        </td>
                                    </tr>
                                    <tr>
                                        <th></th>
                                        <td colspan="2" style="padding:6px">
                                            <ul>
                                                <li><span style="font-size: 95%">Group Policy (<code>gpedit.msc</code>) setting "<code>Turn on Script Execution</code>":</span></li>
                                                <ol>
                                                    <p>
                                                        <table>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><strong>Option</strong></td>
                                                                <td style="padding:6px; font-size: 85%"><strong>PowerShell Equivalent</strong> (concerning all scopes)</td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Not configured</code></td>
                                                                <td style="padding:6px; font-size: 85%">No effect, the default value of this setting</td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Disabled</code></td>
                                                                <td style="padding:6px; font-size: 85%"><code>Restricted</code></td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Enabled</code> - Allow only signed scripts</td>
                                                                <td style="padding:6px; font-size: 85%"><code>AllSigned</code></td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Enabled</code> - Allow local scripts and remote signed scripts</td>
                                                                <td style="padding:6px; font-size: 85%"><code>RemoteSigned</code></td>
                                                            </tr>
                                                            <tr>
                                                                <td style="padding:6px; font-size: 85%"><code>Enabled</code> - Allow all scripts</td>
                                                                <td style="padding:6px; font-size: 85%"><code>Unrestricted</code></td>
                                                            </tr>
                                                        </table>
                                                    </p>
                                                </ol>
                                            </ul>
                                        </td>
                                    </tr>
                                </table>
                            </ul>
                        </p>
                    <p>For more information, please type "<code>Get-ExecutionPolicy -List</code>", "<code>help Set-ExecutionPolicy -Full</code>", "<code>help about_Execution_Policies</code>" or visit <a href="https://technet.microsoft.com/en-us/library/hh849812.aspx">Set-ExecutionPolicy</a> or <a href="http://go.microsoft.com/fwlink/?LinkID=135170">about_Execution_Policies</a>.</p>
                    </li>
                </p>
                <p>
                    <li><code>New-Item -ItemType File -Path C:\Temp\Get-Windows10LockScreenWallpapers.ps1</code><br />
                    Creates an empty ps1-file to the <code>C:\Temp</code> directory. The <code>New-Item</code> cmdlet has an inherent <code>-NoClobber</code> mode built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing file is about to happen. Overwriting a file with the <code>New-Item</code> cmdlet requires using the <code>Force</code>. If the path name and/or the filename includes space characters, please enclose the whole <code>-Path</code> parameter value in quotation marks (single or double):
                        <ol>
                            <br /><code>New-Item -ItemType File -Path "C:\Folder Name\Get-Windows10LockScreenWallpapers.ps1"</code>
                        </ol>
                    <br />For more information, please type "<code>help New-Item -Full</code>".</li>
                </p>
            </ol>
        </td>
    </tr>
</table>




### Contributing

<p>Find a bug? Have a feature request? Here is how you can contribute to this project:</p>

 <table>
   <tr>
      <th><img class="emoji" title="contributing" alt="contributing" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/1f33f.png"></th>
      <td style="padding:6px"><strong>Bugs:</strong></td>
      <td style="padding:6px"><a href="https://github.com/auberginehill/get-windows-10-lock-screen-wallpapers/issues">Submit bugs</a> and help us verify fixes.</td>
   </tr>
   <tr>
      <th rowspan="2"></th>
      <td style="padding:6px"><strong>Feature Requests:</strong></td>
      <td style="padding:6px">Feature request can be submitted by <a href="https://github.com/auberginehill/get-windows-10-lock-screen-wallpapers/issues">creating an Issue</a>.</td>
   </tr>
   <tr>
      <td style="padding:6px"><strong>Edit Source Files:</strong></td>
      <td style="padding:6px"><a href="https://github.com/auberginehill/get-windows-10-lock-screen-wallpapers/pulls">Submit pull requests</a> for bug fixes and features and discuss existing proposals.</td>
   </tr>
 </table>




### www

<table>
    <tr>
        <th><img class="emoji" title="www" alt="www" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/1f310.png"></th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-windows-10-lock-screen-wallpapers">Script Homepage</a></td>
    </tr>
    <tr>
        <th rowspan="39"></th>
        <td style="padding:6px">robledosm: <a href="https://github.com/robledosm/save-lockscreen-as-wallpaper">Save lockscreen as wallpaper</a></td>
    </tr>
    <tr>
        <td style="padding:6px">clayman2: <a href="http://powershell.com/cs/media/p/7476.aspx">Disk Space</a> (or one of the <a href="http://web.archive.org/web/20120304222258/http://powershell.com/cs/media/p/7476.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">lamaar75: <a href="http://powershell.com/cs/forums/t/9685.aspx">Creating a Menu</a> (or one of the <a href="https://web.archive.org/web/20150910111758/http://powershell.com/cs/forums/t/9685.aspx">archive.org versions</a>)</td>
    </tr>
    <tr>
        <td style="padding:6px">Franck Richard: <a href="http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html">Use PowerShell to Remove Metadata and Resize Images</a></td>
    </tr>
    <tr>
        <td style="padding:6px">klumsy: <a href="http://stackoverflow.com/questions/14115415/call-windows-runtime-classes-from-powershell">Call Windows Runtime Classes from PowerShell</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Fred: <a href="https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76ae6430-4993-4422-aa97-8f8ec3ca4e87/selectobject-where?forum=winserverpowershell">select-object | where</a></td>
    </tr>
    <tr>
        <td style="padding:6px">Mike F Robbins: <a href="http://mikefrobbins.com/2015/03/31/powershell-advanced-functions-can-we-build-them-better-with-parameter-validation-yes-we-can/">PowerShell Advanced Functions: Can we build them better?</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://www.askvg.com/windows-10-wallpapers-and-lock-screen-backgrounds/">Download Windows 10 Wallpapers and Lock Screen Backgrounds</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://www.cnet.com/how-to/where-to-find-the-windows-spotlight-photos/">Where to find the Windows Spotlight photos</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://www.windowscentral.com/how-save-all-windows-spotlight-lockscreen-images">How to save Windows Spotlight lockscreen images so you can use them as wallpapers</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/hashhar/Windows-Hacks/blob/master/scheduled-tasks/save-windows-spotlight-lockscreens/save-spotlight.ps1">save-spotlight.ps1</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://answers.microsoft.com/en-us/insider/forum/insider_wintp-insider_personal/windows-10041-windows-spotlight-lock-screen/5b1cddaf-7057-443b-99b6-8c3486a75262">Windows 10041 - Windows Spotlight Lock screen Picture location</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://www.winhelponline.com/blog/find-file-name-lock-screen-image-current-displayed/">How to Find the Current Lock Screen Image File Name and Path in Windows 10?</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://www.tenforums.com/tutorials/38717-windows-spotlight-background-images-find-save-windows-10-a.html">How to Find and Save Windows Spotlight Background Images in Windows 10</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://www.ohmancorp.com/RefWin-windows8-change-pre-login-screen-background.asp">Change the Windows 8 Pre-Login Screen Background</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://social.technet.microsoft.com/Forums/windows/en-US/a8db890c-204f-404a-bf74-3aa4c895b183/cant-customize-lock-or-logon-screen-backgrounds?forum=W8ITProPreRel">Can't customize Lock or Logon screen backgrounds?</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/ff730939.aspx">Adding a Simple Menu to a Windows PowerShell Script</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/ee692804.aspx">The String's the Thing</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string">Powershellv2 - remove last x characters from a string</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://technet.microsoft.com/en-us/library/ee692803.aspx">Working with Hash Tables</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.utility/get-filehash">Get-FileHash</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://nicholasarmstrong.com/2010/02/exif-quick-reference/">EXIF Quick Reference</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/ms630826(v=vs.85).aspx">Windows Image Acquisition (WIA) - Shared Samples</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://kb.winzip.com/kb/entry/207/">WinZip Express for Photos on Windows Server 2003/2008/2012</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/windows/desktop/ms630506(v=vs.85).aspx">ImageFile object</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blogs.msdn.microsoft.com/powershell/2009/03/30/image-manipulation-in-powershell/">Image Manipulation in PowerShell</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://stackoverflow.com/questions/4304821/get-startup-type-of-windows-service-using-powershell">Get startup type of windows service using PowerShell</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.management/get-wmiobject">Get-WmiObject</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://social.microsoft.com/Forums/en-US/4dfe4eec-2b9b-4e6e-a49e-96f5a108c1c8/using-powershell-as-a-photoshop-replacement?forum=Offtopic">Using Powershell as a photoshop replacement</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/system.io.path_methods(v=vs.110).aspx">Path Methods</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://msdn.microsoft.com/en-us/library/ms630826(VS.85).aspx#SharedSample012">Display Detailed Image Information</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://www.experts-exchange.com/questions/25100459/I-need-to-send-the-details-of-a-jpg-file-to-an-array-any-windows-api-to-do-this-or-get-me-started.html">Send the details of a jpg file to an array</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://social.technet.microsoft.com/Forums/windowsserver/en-US/16124c53-4c7f-41f2-9a56-7808198e102a/attribute-seems-to-give-byte-array-how-to-convert-to-string?forum=winserverpowershell">Attribute seems to give byte array. How to convert to string?</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://compgroups.net/comp.databases.ms-access/handy-routine-for-getting-file-metad/1484921">A routine for getting file metadata</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://www.exiv2.org/tags.html">Standard Exif Tags</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blogs.technet.microsoft.com/heyscriptingguy/2013/09/21/powertip-use-powershell-to-send-beep-to-console/">PowerTip: Use PowerShell to Send Beep to Console</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="http://stackoverflow.com/questions/21048650/how-can-i-append-files-using-export-csv-for-powershell-2">How can I append files using export-csv for PowerShell 2?</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/">Remove Unwanted Quotation Marks from CSV Files by Using PowerShell</a></td>
    </tr>
    <tr>
        <td style="padding:6px">ASCII Art: <a href="http://www.figlet.org/">http://www.figlet.org/</a> and <a href="http://www.network-science.de/ascii/">ASCII Art Text Generator</a></td>
    </tr>
</table>




### Related scripts

 <table>
    <tr>
        <th><img class="emoji" title="www" alt="www" height="28" width="28" align="absmiddle" src="https://assets-cdn.github.com/images/icons/emoji/unicode/0023-20e3.png"></th>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/aa812bfa79fa19fbd880b97bdc22e2c1">Disable-Defrag</a></td>
    </tr>
    <tr>
        <th rowspan="26"></th>
        <td style="padding:6px"><a href="https://github.com/auberginehill/firefox-customization-files">Firefox Customization Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ascii-table">Get-AsciiTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-battery-info">Get-BatteryInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-computer-info">Get-ComputerInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-culture-tables">Get-CultureTables</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-directory-size">Get-DirectorySize</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-hash-value">Get-HashValue</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-installed-programs">Get-InstalledPrograms</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-installed-windows-updates">Get-InstalledWindowsUpdates</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-powershell-aliases-table">Get-PowerShellAliasesTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/9c2f26146a0c9d3d1f30ef0395b6e6f5">Get-PowerShellSpecialFolders</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-ram-info">Get-RAMInfo</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/eb07d0c781c09ea868123bf519374ee8">Get-TimeDifference</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-time-zone-table">Get-TimeZoneTable</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/get-unused-drive-letters">Get-UnusedDriveLetters</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/emoji-table">Emoji Table</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/java-update">Java-Update</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/remove-duplicate-files">Remove-DuplicateFiles</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/remove-empty-folders">Remove-EmptyFolders</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/13bb9f56dc0882bf5e85a8f88ccd4610">Remove-EmptyFoldersLite</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://gist.github.com/auberginehill/176774de38ebb3234b633c5fbc6f9e41">Rename-Files</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/rock-paper-scissors">Rock-Paper-Scissors</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/toss-a-coin">Toss-a-Coin</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/unzip-silently">Unzip-Silently</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/update-adobe-flash-player">Update-AdobeFlashPlayer</a></td>
    </tr>
    <tr>
        <td style="padding:6px"><a href="https://github.com/auberginehill/update-mozilla-firefox">Update-MozillaFirefox</a></td>
    </tr>
</table>

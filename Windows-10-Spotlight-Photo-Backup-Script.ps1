 <#
 .SYNOPSIS
    This script copies the beautiful Windows 10 Spotlight Pictures to another folder, so you can use them as background etc. in Windows.
   
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 3.3 Jan 2019 Small update
    Version 3.2 Jan 2019 Small update
    Version 3.1 Mar 2018 Small update and placed code on GitHub
    Version 3.0 May 2016 Update of the copy process, so it skips existing files and other minor fixes
    Version 2.0 March 2016 Redefined the Copy process
	  Version 1.2 March 2016 Updated to count files during copy process
    Version 1.1 March 2016 Fixed bugs
    Version 1.0 March 2016 Initial version
    
	
	.AUTHOR
	Peter Schmidt, Microsoft MVP
    Blog: www.msdigest.net
	
	.CREDITS
	This script is based on the original script work, started by David Rupnik in the Community on: http://www.windowscentral.com/how-save-all-windows-spotlight-lockscreen-images
    Credit to Claus Nielsen - MVP: CLOUD AND DATACENTER (http://xipher.dk/) - for review and input to the copy process.
		
    .DESCRIPTION
	
	Spotlight is a feature in Windows 10 that displays Bing's gorgeous daily images as a slideshow on your lock screen.

	This script copies these beautiful Windows 10 Spotlight pictures to another folder, so you can use them as background etc. in Windows.
	As the images from the original Spotlight locations, is removed and maintained by Windows, so if you really want to backup all, please run this script every day.
	
	The script checks for is the picture is vertical or horizotal and moves to different folders.

	If you want to Schedule it, you can use this command in Task Scheduler:
	powershell.exe -WindowStyle hidden -ExecutionPolicy Bypass [PATH OF YOUR SCRIPT]
    	
	IMPORTANT NOTE: Please set the correct TargetPath to match, where you want you backup located.
	
	.EXAMPLE
    How to run the script: 
    .\Spotlight-Pictures-Backup.ps1
	
    #>
$version = 3.3
#requires -Version 2
$VerbosePreference = "Continue"
#Set the TargetPath to match your wanted settings, in this example it creates a spotlight folder in your OneDrive.
#$TargetPath = "$($env:USERPROFILE)\OneDrive\Spotlight"
$TargetPath = "C:\Spotlight"

$ih=0
$iv=0

Add-Type -AssemblyName System.Drawing
New-Item $TargetPath -ItemType directory -Force
New-Item -Path $TargetPath"\CopyAssets" -ItemType directory -Force
New-Item -Path $TargetPath"\Horizontal" -ItemType directory -Force
New-Item -Path $TargetPath"\Vertical" -ItemType directory -Force

#The location of the Windows Spotlight pictures
foreach($file in (Get-Item -Path "$($env:LOCALAPPDATA)\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets\*"))
{
    #Only files above 100k are copied, which leaves the Ad junk images back.
    if ((Get-Item $file).length -lt 100kb) 
    {
        continue
    }
    Copy-Item -Path $file.FullName -Destination $TargetPath"\CopyAssets\$($file.Name).jpg"
}

foreach($newfile in (Get-Item -Path $TargetPath"\CopyAssets\*"))
{
    $image = New-Object -ComObject WIA.ImageFile
    $image.LoadFile($newfile.FullName)
    #Horizontal pictures are moved to the horizontal folder
    if($image.Width.ToString() -eq '1920')
    {
        If (-Not( Test-Path -Path  "$TargetPath\Horizontal\$($newfile.Name)")) 
        {
            Write-Verbose -Message "Moving file: $($newfile.Fullname) to $TargetPath\Horizontal"
            Copy-Item -Path $newfile.FullName -Destination $TargetPath"\Horizontal" -Force
            $ih++
        }
        Else 
        {
            Write-Verbose -Message "File: $($newfile.Fullname) allready exist in $TargetPath\Horizontal"
        }
    }
    #Vertical pictures are moved to the vertical folder
    elseif($image.Width.ToString() -eq '1080')
    {
        If (-Not( Test-Path -Path  "$TargetPath\Vertical\$($newfile.Name)")) 
        {
            Write-Verbose -Message "Moving file: $($newfile.Fullname) to $TargetPath\Vertical"
            Copy-Item -Path $newfile.FullName -Destination $TargetPath"\Vertical" -Force
            $iv++
        }
        Else 
        {
            Write-Verbose -Message "File: $($newfile.Fullname) allready exist in $TargetPath\Vertical"
        }
    }
}
Remove-Item $TargetPath"\CopyAssets\*";

Write-Host Total number of Horizontal pictures copied is $ih
Write-Host Total number of Vertical pictures copied is $iv
Read-host -prompt "Press enter to complete"
#clear-host;


################################################################################
# Windows PowerShell configuration
#
# Author: William Rideout
# 5/13/2013
################################################################################

################################################################################
# Variables
################################################################################
$user = $env:username
$hname=$env:computername
$Host.UI.RawUI.WindowTitle = "PowerShell on $hname as $user"

################################################################################
# Functions
################################################################################
function Get-Tools
{
	if(Test-Path W:\tools)
	{
		cd W:\tools
		ls
	}
	else
	{
		throw "Mount point W:\ does not exist"
	}
}

function Edit-File
{
	$file = $args[0]
	if(Test-Path "C:\Users\$user\Dropbox\Tools\Notepad++\notepad++.exe")
    {
        & "C:\Users\$user\Dropbox\Tools\Notepad++\notepad++.exe" $file
    }
	else
	{
		throw "Notepad++ not installed or bad path to executable"
	}
}

function Touch-File
{
    $file = $args[0]
    if($file -eq $null)
    {
        throw "No filename supplied"
    }
    
	# Simply update the date on the file if the file exists, otherwise make a
	# new empty file with the specified name
    if(Test-Path $file)
    {
        (Get-ChildItem $file).LastWriteTime = Get-Date
    }
    else
    {
        New-Item $file -ItemType File
    }
}

function Get-WebObject
{
	if($args.Count -ne 2)
	{
		Write-Output "Usage: Get-WebObject <URL> <file>"
	}
	else
	{
		$url = $args[0]
		$filename = $args[1]
		$client = New-Object System.Net.WebClient
		$client.DownloadFile($url, $filename)
	}
}

################################################################################
# Aliases
################################################################################
Set-Alias touch Touch-File
Set-Alias tools Get-Tools
Set-Alias edit Edit-File
Set-Alias wget Get-WebObject

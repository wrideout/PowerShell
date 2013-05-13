###############################################################################
# Author: William Rideout
# 5/13/2013
###############################################################################

###############################################################################
# Variables
###############################################################################
$user = $env:username

###############################################################################
# Functions
###############################################################################
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
	if(Test-Path "C:\Users\Will_Rideout\Dropbox\Tools\Notepad++\notepad++.exe")
    {
        C:\Users\Will_Rideout\Dropbox\Tools\Notepad++\notepad++.exe $file
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

###############################################################################
# Aliases
###############################################################################
Set-Alias touch Touch-File
Set-Alias tools Get-Tools
Set-Alias edit Edit-File

#**********************************************
# Copy all AD groups from one user to another *
#**********************************************

# Check script is running as admin
#=============================================================
param([switch]$elevated)

function Is-Admin 
{
    # Get the current user, and check if they are an Administrator
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

# User does not have elevated privelages
if ((Is-Admin) -eq $false)
{
    if ($elevated) 
    {
        Write-Host "Script needs to be run as admin!" -ForegroundColor DarkRed
    } 
    else
	{
        # Start new elevated instance of PowerShell, user will be prompted with UAC login dialogue
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
	}
    
    # Close current instance of PowerShell
	exit
}

Write-Host "Running as admin...`n" -ForegroundColor Green
#=============================================================


# USER INPUT: SAM name of users to copy from and to
$sourceUser = Read-Host -Prompt 'Input user to copy from'
$destUser = Read-Host -Prompt 'Input user to copy to'

# Fetch the user to copy groups from, stop script if user isn't found.
try
{
	[void] ( Get-ADUser -Identity $sourceUser -ErrorAction Stop )
}
catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
{ 
	Write-Host "User: $sourceUser does not exist" -ForegroundColor DarkRed
	
	break
}

# Fetch the user to copy groups to, stop script if user isn't found.
try
{
	[void] ( Get-ADUser -Identity $destUser -ErrorAction Stop )
}
catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
{ 
	Write-Host "User: $destUser does not exist" -ForegroundColor DarkRed
	
	break
}

Write-Host "Copying Groups...`n==============================================================================`n" -ForegroundColor Cyan

# Pull AD groups from the source user, and add them to the destination user.
Get-ADUser -Identity $sourceUser -Properties memberof -Verbose -ErrorAction Stop | 
Select-Object -ExpandProperty memberof -Verbose |
Add-ADGroupMember -Members $destUser -PassThru -Verbose -ErrorAction Stop

Write-Host "`n==============================================================================`n" -ForegroundColor Cyan
Write-Host "Done!`n" -ForegroundColor Green
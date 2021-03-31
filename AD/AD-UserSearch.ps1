#**********************************************
# Search users in AD and check account status *
#**********************************************

# Check script is running as admin
#=============================================================
param([switch]$elevated)

# Get the current user, and check if they are an Administrator
Function Is-Admin 
{
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

# User is not admin
if ((Is-Admin) -eq $false)
{
    # If user doesn't have required permissions, present error and exit.
    # Otherwise, exit PowerShell and start new instance prompting user to log in as admin.
    if ($elevated) 
    {
        Write-Host "Script requires Admin permissions!" -ForegroundColor DarkRed
    }
    else
	{
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
	}
    
	exit
}

Write-Host "Running as admin...`n" -ForegroundColor Green
#=============================================================


Import-Module ActiveDirectory

#***********
# Main Menu
#***********
Function Main-Menu
{
	Write-Host "Please select an option...`n=================================`n" -ForegroundColor Cyan

	Write-Host "1 : Search for User (SAM Name)" -ForegroundColor Cyan
	Write-Host "2 : Search for User (Email Address)" -ForegroundColor Cyan
	Write-Host "3 : Search for User (Phone Number)" -ForegroundColor Cyan
	Write-Host "4 : Search for User (Employee ID)" -ForegroundColor Cyan
    Write-Host "Q : Quit Application" -ForegroundColor Cyan
	
	Write-Host "`n=================================`n" -ForegroundColor Cyan
	
    # Main menu loop
    #-------------------------------------
    do
    {
        $valid = $false

        Write-Host "Selection : " -NoNewline -ForegroundColor Green
        $selection = Read-Host

        switch ($selection)
	    {
	    	1 {Search-User-SAM;   $valid = $true}
	    	2 {Search-User-Email; $valid = $true}
	    	3 {Search-User-Phone; $valid = $true}
            4 {Search-User-ID;    $valid = $true}

            Q {Write-Host "Quit" -ForegroundColor Yellow; $valid = $true; continue}
            q {Write-Host "Quit" -ForegroundColor Yellow; $valid = $true; continue}

	    	default {Write-Host "Invalid Selection! Try again.`n" -ForegroundColor DarkRed; $valid = $false}
	    }
    }
    until ($valid -eq $true)
    #-------------------------------------	
}

#**********************************************************
# Wrapper for ternary operator, to make code more readable
#**********************************************************
Function Inline-If($If, $Right, $Wrong)
{
    # E.G. 
    # Inline-If ($test-condition, $output-if-true, $output-if-false)
    if ($If)
    {
        $Right
    }
    else
    {
        $Wrong
    }
}

#***************************************
# Main function for returning user info
#***************************************
Function Get-Info($SamName)
{
    Write-Host "Search Results:" -ForegroundColor Magenta
    Write-Host "=================================" -ForegroundColor Magenta

    try
    {
       $samDetails = Get-ADUser -Identity $SamName -Properties '*'

       # Returns the user's OU, derived from their Distinguished Name field
       $userOU = ($samDetails | Select {$_.DistinguishedName -replace '^.*?,(?=[A-Z]{2}=)'} | Format-Table -HideTableHeaders | Out-String).Trim()
    }
    catch
    {
        Write-Host "Get-Info() Error!`n" -ForegroundColor DarkRed
    	
    	break
    }

    Write-Host "USER DETAILS" -ForegroundColor Gray
    Write-Host "-------------------------------------------"  -NoNewline -ForegroundColor Gray

    Write-Host "`nName                  : " (Inline-If (-not ($samDetails.DisplayName -eq $null)) $samDetails.DisplayName "Not Found!") -NoNewline

    Write-Host "`nRole                  : " (Inline-If (-not ($samDetails.Title -eq $null)) $samDetails.Title "Not Found!") -NoNewline
   
    Write-Host "`nEmployee ID           : " (Inline-If (-not ($samDetails.EmployeeID -eq $null)) $samDetails.EmployeeID "Not Found!") -NoNewline

    Write-Host "`nSAM Name              : " (Inline-If (-not ($samDetails.SamAccountName -eq $null)) $samDetails.SamAccountName "Not Found!") -NoNewline
                                        
    Write-Host "`nEmail                 : " (Inline-If (-not ($samDetails.EmailAddress -eq $null)) $samDetails.EmailAddress "Not Found!") -NoNewline
                                        
    Write-Host "`nPhone Number          : " (Inline-If (-not ($samDetails.TelephoneNumber -eq $null)) $samDetails.TelephoneNumber "Not Found!") -NoNewline

    Write-Host "`nOU                    : " (Inline-If (-not ($userOU -eq $null)) $userOU "Not Found!") -NoNewline


    Write-Host "`n`n`nPASSWORD DETAILS" -ForegroundColor Gray
    Write-Host "-------------------------------------------" -NoNewline -ForegroundColor Gray

    Write-Host "`nLast Logon            : " (Inline-If (-not ($samDetails.LastLogonDate -eq $null)) $samDetails.LastLogonDate "Never") -NoNewline

    Write-Host "`nP/W Last Set          : " (Inline-If (-not ($samDetails.PasswordLastSet -eq $null)) $samDetails.PasswordLastSet "Never Set!") -NoNewline

    Write-Host "`nLast P/W Incorrect    : " (Inline-If (-not ($samDetails.LastBadPasswordAttempt -eq $null)) $samDetails.LastBadPasswordAttempt "Never") -NoNewline

    Write-Host "`nP/W Expired?          : " (Inline-If ($samDetails.PasswordExpired -eq  $false) "Password OK!" "Password Expired!") -NoNewline

    Write-Host "`nLocked Out?           : " (Inline-If ($samDetails.LockedOut -eq  $false) "Not Locked Out" "Locked Out!") -NoNewline
    
    
    Write-Host "`n`n`nACCOUNT DETAILS" -ForegroundColor Gray
    Write-Host "-------------------------------------------"  -NoNewline -ForegroundColor Gray

    Write-Host "`nAccount Created       : " (Inline-If (-not ($samDetails.WhenCreated -eq $null)) $samDetails.WhenCreated "Not Found!") -NoNewline

    Write-Host "`nAccount Last Modified : " (Inline-If (-not ($samDetails.WhenCreated -eq $null)) $samDetails.WhenCreated "Not Found!") -NoNewline

    Write-Host "`nAccount Enabled?      : " (Inline-If ($samDetails.Enabled -eq  $true) "Account Enabled" "Account Disabled!") -NoNewline
    
    Write-Host "`n`n=================================`n" -ForegroundColor Magenta
}

#*****************************************************
# Wrapper for calling Get-Info with inputted SAM name
#*****************************************************
Function Search-User-SAM
{
    Write-Host "`n=================================`n" -ForegroundColor Cyan

    Write-Host "Enter SAM name : " -NoNewline -ForegroundColor Green
    $samName = Read-Host

    Write-Host # Spacing

    # Error checking
    #-------------------------
    try
    {
    	[void] ( Get-ADUser -Identity $samName -ErrorAction Stop )
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
    { 
    	Write-Host "User: $samName does not exist!`n" -ForegroundColor DarkRed
    	
    	break
    }
    catch [Microsoft.ActiveDirectory.Management.ADServerDownException]
    { 
    	Write-Host "Can't find server with AD services running!`n" -ForegroundColor DarkRed
    	
    	break
    }
    catch
    {
        Write-Host "$samName does not exist!`n" -ForegroundColor DarkRed
    	
    	break
    }
    #-------------------------

    Get-Info($samName)
}

#**********************************************************
# Wrapper for calling Get-Info with inputted email address
#**********************************************************
Function Search-User-Email
{
    Write-Host "`n=================================`n" -ForegroundColor Cyan

    Write-Host "Enter Email Address: " -NoNewline -ForegroundColor Green
    $email = Read-Host

    Write-Host # Spacing

    # Error checking
    #-------------------------
    try
    {
    	[void] ( Get-ADUser -Filter "EmailAddress -like '*$email*'" -ErrorAction Stop )
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
    { 
    	Write-Host "$email does not exist!`n" -ForegroundColor DarkRed
    	
    	break
    }
    catch [Microsoft.ActiveDirectory.Management.ADServerDownException]
    { 
    	Write-Host "Can't find server with AD services running!`n" -ForegroundColor DarkRed
    	
    	break
    }
    catch
    {
        Write-Host "$email does not exist!!`n" -ForegroundColor DarkRed
    	
    	break
    }
    #-------------------------

    $samName = Get-ADUser -Filter "EmailAddress -eq '$email'" -Properties SamAccountName | Select-Object SamAccountName | % {$_.SamAccountName}

    Get-Info($samName)
}

#*********************************************************
# Wrapper for calling Get-Info with inputted phone number
#*********************************************************
Function Search-User-Phone
{
    # TODO: Check phone number for correct length
    # TODO: Check common variations of phone number (i.e. 07525000000 vs 07525 000000)

    Write-Host "`n=================================`n" -ForegroundColor Cyan

    Write-Host "Enter phone number : " -NoNewline -ForegroundColor Green
    $phoneNum = Read-Host

    # Common formatting variations
    $phoneNumSplit = $phoneNum -replace '[^0-9]'  #07525000000
    $phoneNumSplit = $phoneNumSplit.Insert(5,' ') #07525 000000
    $phoneNumSplit2 = $phoneNumSplit.Insert(9,' ') #07525 000 000

    Write-Host "`n-------------------------" -ForegroundColor Gray
    Write-Host "Searching..." -ForegroundColor Gray

    # Error checking
    #-------------------------
    try
    {
    	[void] ( Get-ADUser -Filter "TelephoneNumber -like '*$phoneNum*'" -ErrorAction Stop )
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
    { 
    	Write-Host "$phoneNum does not exist!`n" -ForegroundColor DarkRed
    	
    	#break
    }
    catch [Microsoft.ActiveDirectory.Management.ADServerDownException]
    { 
    	Write-Host "Can't find server with AD services running!`n" -ForegroundColor DarkRed
    	
    	#break
    }
    #-------------------------

    Write-Host "$phoneNum : " -ForegroundColor Gray -NoNewline

    if ((Get-ADUser -Filter "TelephoneNumber -like '*$phoneNum*'") -eq $null)
    {
        Write-Host "NOT FOUND" -ForegroundColor Gray
        $phoneNum = $phoneNumSplit
        Write-Host "Searching..." -ForegroundColor Gray
        Write-Host "$phoneNum : " -ForegroundColor Gray -NoNewline
    }

    if ((Get-ADUser -Filter "TelephoneNumber -like '*$phoneNum*'") -eq $null)
    {
        Write-Host "NOT FOUND" -ForegroundColor Gray
        $phoneNum = $phoneNumSplit2
        Write-Host "Searching..." -ForegroundColor Gray
        Write-Host "$phoneNum : " -ForegroundColor Gray -NoNewline
    }

    if (-not ((Get-ADUser -Filter "TelephoneNumber -like '*$phoneNum*'") -eq $null))
    {
        Write-Host "FOUND" -ForegroundColor Gray
    }
    else
    {
        Write-Host "NOT FOUND" -ForegroundColor Gray
        
    }
    Write-Host "-------------------------`n" -ForegroundColor Gray

    if (-not ((Get-ADUser -Filter "TelephoneNumber -like '*$phoneNum*'") -eq $null))
    {
        $samName = Get-ADUser -Filter "TelephoneNumber -like '*$phoneNum*'" -Properties SamAccountName | Select-Object SamAccountName | % {$_.SamAccountName}
    
        Get-Info($samName)
    }
    
}

#********************************************************
# Wrapper for calling Get-Info with inputted employee ID
#********************************************************
Function Search-User-ID
{
    Write-Host "`n=================================`n" -ForegroundColor Cyan

    Write-Host "Enter employee ID : " -NoNewline -ForegroundColor Green
    $empID = Read-host

    Write-Host # Spacing

    # Error checking
    #-------------------------
    try
    {
    	[void] ( Get-ADUser -Filter "EmployeeID -like '*$empID*'" -ErrorAction Stop )
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
    { 
    	Write-Host "$empID does not exist!`n" -ForegroundColor DarkRed
    	
    	break
    }
    catch [Microsoft.ActiveDirectory.Management.ADServerDownException]
    { 
    	Write-Host "Can't find server with AD services running!`n" -ForegroundColor DarkRed
    	
    	break
    }
    catch
    {
        Write-Host "$empID does not exist!!`n" -ForegroundColor DarkRed
    	
    	break
    }
    #-------------------------

    $samName = Get-ADUser -Filter "EmployeeID -eq '$empID'" -Properties SamAccountName | Select-Object SamAccountName | % {$_.SamAccountName}

    Get-Info($samName)
}

#***********************
# Main Application Loop
#***********************
do
{
    Write-Host "-----------------------------------------------------------------------------`n" -ForegroundColor DarkGray
    
    Main-Menu
    
    Write-Host "-----------------------------------------------------------------------------`n" -ForegroundColor DarkGray

    Write-Host "`nReturn to Menu? " -NoNewline -ForegroundColor Green
    Write-Host "[y:Return to Menu] [n:Exit Application] : " -NoNewline -ForegroundColor Gray
    $exit = Read-Host

    switch ($exit)
	{
		y {Write-Host "Returning to Menu.`n`n"  -ForegroundColor Gray;   $valid = $true;   $exit = $false;  continue}
        Y {Write-Host "Returning to Menu.`n`n"  -ForegroundColor Gray;   $valid = $true;   $exit = $false;  continue}
		
        n {Write-Host "Goodbye...`n"            -ForegroundColor Gray;   $valid = $true;   $exit = $true;   continue}
        N {Write-Host "Goodbye...`n"            -ForegroundColor Gray;   $valid = $true;   $exit = $true;   continue}
        
		default {Write-Host "Invalid Selection! Try again.`n" -ForegroundColor DarkRed; $valid = $true; $exit = $false}
	}
}
until (($exit -eq $true) -and ($valid -eq $true))
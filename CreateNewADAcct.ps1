<# Creates new account information in AD using information from a CSV file.

    .SYNOPSIS
        This powershell script will parse a CSV file and create AD Accounts along with their respective email mailboxes.

    .USAGE
        Ensure we have the files to parse through, will loop over the given csv and create AD and Email information as needed.

    .PARAMETER newaccountCSV (required)
        The accounts file.

    .PARAMETER exchangeserver (required if not set in configrc)
        The exchange server.

    .PARAMETER password (required if not set in configrc)
        The password to set the new accounts to.

    .PARAMETER emailDomain (required if not set in configrc)
        The domain for your email information.

    .PARAMETER configFile (optional)
        The path to the config file, if not set it will look in the same place the script is located.

    .PARAMETER successLog (optional)
        The path to store the success log, if not set it will set in same location as script, and be named SuccessLog.txt

    .PARAMETER errorLog (optional)
        The path to store the error log, if not set it will set in same location as script, and be named ErrorLog.txt

#>

[CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [string]$newaccountCSV,

        [Parameter()]
        [string]$exchangeserver,

        [Parameter()]
        [string]$password,

        [Parameter()]
        [string]$emailDomain,

        [Parameter()]
        [string]$configFile,

        [Parameter()]
        [string]$successLog,

        [Parameter()]
        [string]$errorLog
    )

# Set errors to silent if we are not in verbose mode.
if ($PSCmdlet.Myinvocation.BoundParameters["Verbose"].IsPresent) {
    $ErrorActionPreference = "SilentlyContinue"
}

# Set up default if configFile is not already set.
if (!$configFile) {
    $configFile = ".\configrc.ps1"
}

# Source our config file
. $configFile

if (!$exchangeserver) {
    Write-Error "No -exchangeserver passed."
    return
}

if (!$password) {
    Write-Error "No -password passed."
    return
}

if (!$emailDomain) {
    Write-Error "No -emailDomain passed."
    return
}

if (!$successLog) {
    $successLog = ".\SuccessLog.txt"
}

if (!$errorLog) {
    $errorLog = ".\ErrorLog.txt"
}

# Build the connection to the exchange server powershell information.
$connectionURI = "http://" + $exchangeserver + "/powershell"

# Get the current date for setting up logging.
$date = Get-Date

# Instantiate our Log file information:
Write-Verbose "Instantiating our Log Files for this session"
Add-Content $successLog "-----------------------------------------------------------------"
Add-Content $successLog $date
Add-Content $successLog "-----------------------------------------------------------------"

Add-Content $errorLog "-----------------------------------------------------------------"
Add-Content $errorLog $date
Add-Content $errorLog "-----------------------------------------------------------------"

# Exchange Params
$sessionprms = @{
    ConfigurationName = "Microsoft.Exchange"
    ConnectionUri = $connectionURI
    Authentication = "kerberos"
}

# Establishing Connection to Exchange
Write-Verbose "Establishing connection to Exchange"
$session = New-PSSession @sessionprms

# Import Exchange Information
Write-Verbose "Importing Exchange Session information for this session"
Import-PSSession -Session $session

# Import Active Directory Information.
Write-Verbose "Importing ActiveDirectory for this session"
Import-Module ActiveDirectory 

# Import CSV
Write-Verbose "Loading the CSV: $newaccountCSV"
$csv = @()
$csv = Import-Csv -Delimiter "," -Path $newaccountCSV

# Getting search base DN from AD
Write-Verbose "Getting search base DN of active directory"
$searchbase = Get-ADDomain | ForEach {$_.DistinguishedName}

#Loop through all items in the CSV  
ForEach ($user In $csv) {
    # Separate information to component variables
    $lastname = $user.Lastname
    $firstname = $user.Firstname
    $title = $user.JobTitle
    $department = $user.Department
    $office = $user.Office
    $manager = $user.ManagerEmail

    # Create the username we will be using along with our display name and UPN.
    $username = $firstname.Substring(0,1).tolower() + $lastname.tolower()
    $displayName = $lastname + ", " + $firstname
    $upn = $username + "@" + $emailDomain

    # Finding manager
    $manager = Get-ADUser -Filter "(emailaddress -like '$manager')" | SELECT -First '1' -ExpandProperty SamAccountName

    # Default AD Groups to Join during creation. (Format: "ABC","DEF"; Not: "ABC,DEF")
    $group= "All E-Mail Users"
    
    # Check if the User exists
    $userexist = Get-ADUser -LDAPFilter "(samAccountName=$username)"

    # If the user does exist, output to error log and describe why, continue to next user.
    if ($userexist) {
        Write-Verbose "User Already Exists: $username"
        Add-Content $ErrorLog "-------------------------------------------------------------------"
        Add-Content $ErrorLog "User already exists: $username"
        Add-Content $ErrorLog "-------------------------------------------------------------------"
        continue
    }

    # Create the user parameters
    $createprms = @{
        Name = $displayName
        SamAccountName = $username
        UserPrincipalName = $upn
        DisplayName = $displayName
        GivenName = $firstname
        SurName = $lastname
        AccountPassword = (ConvertTo-SecureString $password -AsPlainText -Force)
        Enabled = $True
        ChangePasswordAtLogon = $changepwatlogon
        Title = $title
        Description = $title
        EmailAddress = $upn
    }

    if ($ou) {
        $createprms.Path = $ou
    }
    if ($homedriveltr) {
        $createprms.HomeDrive = $homedriveltr+':'
    }
    if ($homedrive) {
        $createprms.HomeDirectory = $homedrive
    }
    if ($logonscript) {
        $createprms.ScriptPath = $logonscript
    }
    if ($company) {
        $createprms.Company = $company
    }
    if ($logonscript) {
        $createprms.ScriptPath = $logonscript
    }
    if ($profilePath) {
        $createprms.ProfilePath = $profilePath
    }
    if ($office) {
        $createprms.Office = $office
    }
    if ($department) {
        $createprms.Department = $department
    }
    if ($manager) {
        $createprms.Manager = $manager
    }

    # Create the user account
    Write-Verbose "Creating User with account name: $username"
    $create = New-ADUser @createprms

    # Another way to set manager shouldn't be needed, but is here just in case.
    #Set-AdUser "$SAM" -Manager $Manager

    # If the user failed to create, log it
    # Check if the User exists
    $userexist = Get-ADUser -LDAPFilter "(samAccountName=$username)"
    if (!$userexist) {
        Write-Verbose "User failed to create: $username"
        Add-Content $ErrorLog "-------------------------------------------------------------------"
        Add-Content $ErrorLog "User create failed: $username"
        Add-Content $ErrorLog "-------------------------------------------------------------------"
        continue
    }

    # User Created!
    Write-Verbose "User Successfully Created: $username"
    Add-Content $SuccessLog "-------------------------------------------------------------------"
    Add-Content $SuccessLog "User created: $username"
    Add-Content $SuccessLog "-------------------------------------------------------------------"

    # Wait 5 seconds for new account to propagate
    Start-Sleep -s 5

    # Adding to groups
    Write-Verbose "Assigning user to default groups"
    $groupADD = Add-ADPrincipalGroupMembership -Identity $username -MemberOf $group

    # Group Join failed, log it.
    #if (!$groupADD) {
    #    Write-Verbose "Group join failed for user: $username"
    #    Add-Content $ErrorLog "-------------------------------------------------------------------"
    #    Add-Content $ErrorLog "Group join failed for user: $username"
    #    Add-Content $ErrorLog "-------------------------------------------------------------------"
    #} else {
    #    Write-Verbose "Group(s) joined successfully for user: $username"
    #    Add-Content $SuccessLog "-------------------------------------------------------------------"
    #    Add-Content $SuccessLog "Group(s) joined successfully for: $username"
    #    Add-Content $SuccessLog "-------------------------------------------------------------------"
    #}

    # Mailbox Params
    $mailprm = @{
        Identity = $username
        Alias = $username
    }
    if ($maildatabase) {
        $mailprm.Database = $maildatabase
    }

    # Wait 5 seconds to ensure Account information propagates for Exchange to see it.
    Start-Sleep -s 5

    # Creating Mailbox on Exchange 2010
    Write-Verbose "Enabling/Creating E-mail on Exchange server"
    Enable-Mailbox @mailprm
    
    # If mailbox enable/create fails, log it and continue
    #if (!$mailbox) {
    #    Write-Verbose "Mailbox create/enable failed for user: $username"
    #    Add-Content $ErrorLog "-------------------------------------------------------------------"
    #    Add-Content $ErrorLog "Email create/enable failed for user: $username"
    #    Add-Content $ErrorLog "-------------------------------------------------------------------"
    #    continue;
    #}

    # All good, log success and move on.
    Write-Verbose "Mailbox created/enabled successfully: $username"
    Add-Content $SuccessLog "-------------------------------------------------------------------"
    Add-Content $SuccessLog "Mailbox created/enabled successfully for: $username"
    Add-Content $SuccessLog "-------------------------------------------------------------------"
}
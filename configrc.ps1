<# General Configuration File For Account Creation in AD
    .SYNOPSIS
        This powershell script will be the Configuration File for our Creation of AD Accounts/Mailbox information.
        This is simply to generalize the New AD Account script so others may use it in their own environments.

    .USAGE
        Required Variables include:
        @param $Password string The password to assign the new accounts.
        @param $emailDomain string The domain name to use for creating email. (Becomes part of the UPN later too)
        @param $exchangeserver string the exchange server name.
        
        Optional Variables include:
        @param $Company string The company to associate this too.
        @param $OU string This is the OU Placement String. This is the full OU path e.g. "OU=New Users,CN=Users,DC=contoso,DC=com
        @param $homedrive string this is the path (network share too) for the users homedrive
        @param $profilePath string this is the path to setup the users profile information.
        @param $logonscript string this is the path/script name for when users logon.
        @param $maildatabase string this is the Mailbox database, should not be required.
        @param $homedriveltr string this is the home drive letter to use.
        @param $changepwatlogon bool this is if to reset pw at first logon or not.

        Logging file information (optional)
        @param $successLog string The path to write success information to.
        @param $errorLog string The path to write error information to.
#>

## Required Variables:
$password = 'S0m3P@$$word!'
$emailDomain = "example.notadomain"
$exchangeserver = "someexchange.example.notadomain"

## Optional Variables:
$company = "Example"
$ou = "OU=New Users,CN=Users,DC=example,DC=notadomain"
$homedrive = '\\someuserdriveshare\users$\%username%'
$profilePath = "\\someuserprofileshare\Profiles\%username%"
$logonscript = "Somescriptnamehere.vbs"
$maildatabase = "SomeMailBox Database name here"
$homedriveltr = "U"
$changepwatlogon = $False

## Logging setup (optional)
$successLog = ".\SuccessLog.txt"
$errorLog = ".\ErrorLog.txt"

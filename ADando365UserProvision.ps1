
#Set your variables here#

$Company = "Your Company Name here"
$OU = "Enter DN of OU where the user is created here"
$Department = "Depart of the user here"
$DomainName = "Domain Name used for email and UPN"
$ADConnectServer = "Azure AD Connect Server Name for Remote PS"

if (!(Get-Module -Name Msonline))
{
    Import-Module Msonline
}else
{
    Write-Host -ForegroundColor Green `nMsonline PS Module is already loaded.
}

if (!($Credential))
{
    $Credential = Get-Credential
}else
{
    Write-Host -ForegroundColor Green `nThe Variable '$Credential' already contains a creadential. `nTo reset Credential, use '$Credential = $null'  to clear or '$Credential = Get-Credential' command to Re-Enter Credential
}

if (!(Get-MsolDomain -ErrorAction Ignore))
{
    Connect-MsolService -Credential $Credential
}

$ExSession = Get-PSSession | Where-Object `
{($_.ComputerName -eq "outlook.office365.com") -and `
 ($_.State -eq "Opened") -and `
 ($_.ConfigurationName -eq "Microsoft.Exchange")} 
if (!($ExSession))
{
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
    Import-PSSession $Session
}else
{
    Write-Host -ForegroundColor Green `nWrite-Host Exchange Online PSSession Already Exisit. `nIf You Wish to Reconnect, Use Remove-PSSession '$Session' command and then Re-Run the Script
}




Write-Host `n

$Licences = Get-MsolAccountSku

$Licence_Count = 0

$Unavailble_Licence = New-Object System.Collections.ArrayList

Foreach ($Licence in $Licences)
{
    $Licence_Count++
    
    $Licence_Remain = $Licence.ActiveUnits - $Licence.ConsumedUnits

    Write-Host $Licence_Count : $Licence.AccountSkuId "(" $Licence_Remain "More Remaining )"

    if ($Licence_Remain -eq 0)
    {
        $null = $Unavailble_Licence.Add($Licence_Count)
    }
}

[int]$global:Licence_No = Read-Host `nSelect Licence Option Number

while (!(($Licence_No -ge 1) -and ($Licence_No -le $Licence_Count)) -or ($Unavailble_Licence -contains $Licence_No))
{
   
    Write-Host -ForegroundColor Red `nInvalid Selection.`n
    
    if ($Unavailble_Licence -contains $Licence_No)
    {
        Write-Host -ForegroundColor Red The Licence selected does not have a remaining Unit. Please select another Licence. `n
    }
    $Licence_No = Read-Host `n Select Licence Option Number, agian.
}

$Licence_No = $Licence_No - 1

$Firstname = Read-Host -Prompt "Enter First Name"
$Firstname = $Firstname.substring(0,1).toupper() + $FirstName.substring(1)

$LastName = Read-Host -Prompt "Enter Last Name" 
$LastName = $LastName.substring(0,1).toupper() + $LastName.substring(1)

$FullName = $Firstname + " "+  $LastName

$EmailAlias = Read-Host -Prompt "Enter Email Alias."
$EmailAlias = $EmailAlias.substring(0).tolower()
$EmailAddress = $EmailAlias + '@' + $DomainName 
$DN = "CN=" + $FullName + "," + $OU
$Password = Read-Host -AsSecureString -Prompt "Enter the Password"

New-ADUser -Company $Company -Department $Department -DisplayName $FullName -EmailAddress $EmailAddress `
 -GivenName $Firstname -Name $FullName -Path $OU -SamAccountName $EmailAlias -Surname $LastName `
 -Type "user" -UserPrincipalName  $EmailAddress  

Set-ADAccountPassword $DN -NewPassword $Password -Reset:$true 

Set-ADUser -Identity $EmailAlias -Add @{Proxyaddresses="SMTP:"+$EmailAddress}

Enable-ADAccount $EmailAlias 

Write-Host -ForegroundColor Green The User Account Created
Write-Host -ForegroundColor Green Full Name is $FullName
Write-Host -ForegroundColor Green Email Address is $EmailAddress
Write-Host -ForegroundColor Green Azure AD Connect Server is $ADConnectServer

Start-Sleep 10

Invoke-Command -ComputerName $ADConnectServer -ScriptBlock {Import-Module ADSync; Start-ADSyncSyncCycle -PolicyType Delta}

$GetUser = Get-MsolUser -UserPrincipalName $EmailAddress -ErrorAction SilentlyContinue

while (!($GetUser))
{
    $GetUser = Get-MsolUser -UserPrincipalName $EmailAddress -ErrorAction SilentlyContinue
    Write-Host "Waiting for the user account to be synchronised."
    Start-Sleep 10
}

Write-Host -ForegroundColor Green "`nThe User Account is synchronised."

Set-MsolUser -UserPrincipalName $EmailAddress -UsageLocation GB
Set-MsolUserLicense -UserPrincipalName $EmailAddress -AddLicenses $Licences[$Licence_No].AccountSkuId

Write-Host -ForegroundColor Green "`nThe User Account has been provisioned successully."
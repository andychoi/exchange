

################
#  Parameters  #
################
param(
    [ValidateSet("export", "import", "delete")]
    [string]$operation,
    [switch]$ConfirmDeletes
)



###############
#  Variables  #
###############

Write-Host "Enter the Exchange Online Credentials" -ForegroundColor Yellow
$outputcsv = ".\galSyncc.csv"
$inputcsv = ".\galSyncc-test.csv"
#$RecipientTypes = ('UserMailbox', 'MailUser', 'MailContact', 'MailUniversalSecurityGroup', 'MailUniversalDistributionGroup', 'RemoteUserMailbox')
$RecipientTypes = ('UserMailbox')

$DomainsToExclude = ("domain1.com", "domain2.com")  #hke.local

#$galTemp, $galTempp, $galTempfinal = [System.Collections.ArrayList]@()
$galTemp, $galTempp, $galTempfinal = @()

$unlimited = "unlimited"  #30
#Note: The ConnectionUri value is http, not https
#$onpremisesConnectionUri = "http://irvhisechp10.hke.local/PowerShell"
$onpremisesConnectionUri = "http://owa.hisna.com/PowerShell"
$EXOConnectionUri = "https://outlook.office365.com/powerShell-liveID?serializationLevel=Fullset-ads"
$SyncAttribute = "CustomAttribute10"
$SyncAttributeValue = "Sync"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$logfi = $DesktopPath + "\GalLog_" + [DateTime]::Now.ToString("yyyyMMdd") + ".csv"
if (!(Test-Path($logfi))) { "Date,Status,Message, PrimarySmtpAddress,DisplayName">$logfi }
    

Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $False)]
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
        [String]
        $Level = "INFO",

        [Parameter(Mandatory = $True)]
        [string]
        $Message,

        [Parameter(Mandatory = $False)]
        [string]
        $logfile
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = $Stamp + "," + $Level + "," + $Message
    If ($logfile) {
        
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}


####################
#  Authentication  #
####################

Function auth_OnPremises {
    $ses = ""
    $ses = (Get-PSSession).Name
    if ($ses -ne $null) {
        Remove-PSSession -Session (Get-PSSession) -ErrorAction SilentlyContinue
    }
    #To delete all the PSSessions in the current session, type Get-PSSession | Remove-PSSession

    #Write-Host "Enter Onpremises credentials:" -ForegroundColor Red -BackgroundColor Yellow
#modified by Andy:  https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-servers-using-remote-powershell?view=exchange-ps
#   if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1) {
#       . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
#       Connect-ExchangeServer -auto -AllowClobber
#       #cls
#   } else {
#       Write-Warning "Exchange Server management tools are not installed on this computer."
#       EXIT
#   }
    $OnPremisesCredential = Get-Credential -Message "Exchange on-premises Credential"
    #   $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $onpremisesConnectionUri -Credential $OnPremisesCredential -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $onpremisesConnectionUri -Authentication Kerberos -AllowRedirection 
    if(-not($Session)) {
        Write-Warning "Exchange Server login error"
        EXIT
    }
    
    Import-PSSession $Session -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking -AllowClobber | Out-Null
    Write-Log -Message "Connected to on premises Exchange,," -Level INFO -logfile $logfi
    $Input = Read-Host -Prompt "Connected to on premises Exchange. Press any key to continue"
}

Function auth_Online {
    $ses = ""
    $ses = (Get-PSSession).Name
    if ($ses -ne $null) {
        Remove-PSSession -Session (Get-PSSession) -ErrorAction SilentlyContinue
    }

    ##To connect exchange online with MFA, you need to install Microsoft's Exchange Online PowerShell Module(EXO).
    try {
        $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)

        #Importing Exchange MFA Module 
        If ($MFAExchangeModule) {
            . "$MFAExchangeModule"
        }
       # Connect-EXOPSSession  
        Connect-ExchangeOnline
    }
    catch {
        Write-Host "We couldn't authenticate to Exchange Online" -ForegroundColor Red -BackgroundColor Yellow
        exit
    }
    $Input = Read-Host -Prompt "Connected to on Exchange Online. Press any key to continue"
}
####################################
#  Configure the excluded domains  #
####################################
function get_Excluded_Domains() {
    Write-Host "We need to authenticate to EXO to get the accepted domains" -ForegroundColor black -BackgroundColor Yellow
    #temp for testing purpose comment out
    #auth_Online
    if ($DomainsToExclude) {
        $domainsexcluded = (Get-AcceptedDomain -ErrorAction Stop).DomainName | ? { $_ -notlike "*onmicrosoft.com" }
        $domainsexcluded += $DomainsToExclude
        $script:regexdomain = ($domainsexcluded | % { $_ }) -join "|"
        Write-Host "These domains going to be excluded: $regexdomain" -ForegroundColor Yellow
        Write-Log -Message "Domains excluded: $regexdomain ,," -Level INFO -logfile $logfi
    }
    else {
        $domainsexcluded = (Get-AcceptedDomain -ErrorAction Stop).DomainName | ? { $_ -notlike "*onmicrosoft.com" }
        $script:regexdomain = ($domainsexcluded | % { $_ }) -join "|"
        Write-Host "These domains going to be excluded: $regexdomain" -ForegroundColor Yellow
        Write-Log -Message "Domains excluded: $regexdomain ,," -Level INFO -logfile $logfi
    }
    
}

######################################
#  Functions to add user properties  #
######################################

function adduserprop($user) {
    $returnuser = get-user ($user.guid).ToString() | select Phone, MobilePhone, Company, Title, Department, Office, FirstName, LastName
    return $returnuser
}


function finaluserprop($galTemp) {
    $galTempfinal = [System.Collections.ArrayList]@()
    foreach ($gal in $galTemp) {
        $userprop = adduserprop $gal.guid
        $gal.Phone = $userprop.Phone
        $gal.MobilePhone = $userprop.MobilePhone
        $gal.Company = $userprop.Company
        $gal.Title = $userprop.Title
        $gal.Department = $userprop.Department
        $gal.Office = $userprop.Office
        $gal.FirstName = $userprop.FirstName
        $gal.LastName = $userprop.LastName
        $galTempfinal += $gal
    }
    return $galTempfinal
}

######################################
#  Exporting the data to a csv file  #
######################################

function export() { 
    ###############################
    #  Exporting All receipients  #
    ###############################
    Write-Host "We need to authenticate to Exchange on premises to export all the recipients" -ForegroundColor black -BackgroundColor Yellow
    if (Test-Path($outputcsv)) { Clear-Content $outputcsv }
    auth_OnPremises
    Set-AdServerSettings -ViewEntireForest $True
    Write-Host "Exporting all Recipients" -ForegroundColor Green
    $c = 0
    $rec_count = $RecipientTypes.Count
    ForEach ($type in $RecipientTypes) {
        $c++
        Write-Progress -Activity "$c out of $rec_count" -Status $type -PercentComplete (($c / $rec_count) * 100)
         
        Switch ($type) {
            UserMailbox {
                Write-Host "    Exporting Mailboxes" -ForegroundColor Yellow
                Set-AdServerSettings -ViewEntireForest $True
                $galTemp = Get-mailbox -ResultSize $unlimited -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | ? { ($_.Alias -notmatch "DiscoverySearchMailbox") -and ($_.PrimarySmtpAddress -notmatch $regexdomain) -and ($_.OrganizationalUnit -notmatch 'HKEnterprise.*Disabled-Accounts') } | Select Alias, DisplayName, @{n = "EmailAddresses"; e = { $_.EmailAddresses -join ";" } }, ExternalEmailAddress, FirstName, HiddenFromAddressListsEnabled, LastName, LegacyExchangeDn, Name, PrimarySmtpAddress, RecipientType, @{n = "Phone"; e = { "" } }, @{n = "MobilePhone"; e = { "" } }, @{n = "Company"; e = { "" } }, @{n = "Title"; e = { "" } }, @{n = "Department"; e = { "" } }, @{n = "Office"; e = { "" } }, guid
                $galTempp += finaluserprop($galTemp)
            }
            MailUser {
                Write-Host "    Exporting MailUser" -ForegroundColor Yellow
                Set-AdServerSettings -ViewEntireForest $True
                $galTemp = Get-MailUser -Resultsize $unlimited -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | ? { ($_.Alias -notmatch "DiscoverySearchMailbox") -and ($_.PrimarySmtpAddress -notmatch $regexdomain) } | Select Alias, DisplayName, @{n = "EmailAddresses"; e = { $_.EmailAddresses -join ";" } }, ExternalEmailAddress, FirstName, HiddenFromAddressListsEnabled, LastName, LegacyExchangeDn, Name, PrimarySmtpAddress, RecipientType, @{n = "Phone"; e = { "" } }, @{n = "MobilePhone"; e = { "" } }, @{n = "Company"; e = { "" } }, @{n = "Title"; e = { "" } }, @{n = "Department"; e = { "" } }, @{n = "Office"; e = { "" } }, guid
                $galTempp += finaluserprop($galTemp)
            }
            MailContact {
                Write-Host "    Exporting MailContacts" -ForegroundColor Yellow
                Set-AdServerSettings -ViewEntireForest $True
                $galTempp += Get-MailContact -ResultSize $unlimited -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | ? { ($_.Alias -notmatch "DiscoverySearchMailbox") -and ($_.PrimarySmtpAddress -notmatch $regexdomain) } | Select Alias, DisplayName, @{n = "EmailAddresses"; e = { $_.EmailAddresses -join ";" } }, ExternalEmailAddress, FirstName, HiddenFromAddressListsEnabled, LastName, LegacyExchangeDn, Name, PrimarySmtpAddress, RecipientType, @{n = "Phone"; e = { "" } }, @{n = "MobilePhone"; e = { "" } }, @{n = "Company"; e = { "" } }, @{n = "Title"; e = { "" } }, @{n = "Department"; e = { "" } }, @{n = "Office"; e = { "" } }, guid
            }
            MailUniversalSecurityGroup {
                Write-Host "    Exporting MailUniversalSecurityGroup" -ForegroundColor Yellow
                Set-AdServerSettings -ViewEntireForest $True
                $galTempp += Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize $unlimited -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | ? { ($_.Alias -notmatch "DiscoverySearchMailbox") -and ($_.PrimarySmtpAddress -notmatch $regexdomain) } | Select Alias, DisplayName, @{n = "EmailAddresses"; e = { $_.EmailAddresses -join ";" } }, ExternalEmailAddress, FirstName, HiddenFromAddressListsEnabled, LastName, LegacyExchangeDn, Name, PrimarySmtpAddress, RecipientType, @{n = "Phone"; e = { "" } }, @{n = "MobilePhone"; e = { "" } }, @{n = "Company"; e = { "" } }, @{n = "Title"; e = { "" } }, @{n = "Department"; e = { "" } }, @{n = "Office"; e = { "" } }, guid
            }
            MailUniversalDistributionGroup {
                Write-Host "    Exporting MailUniversalDistributionGroup" -ForegroundColor Yellow
                Set-AdServerSettings -ViewEntireForest $True
                $galTempp += Get-DistributionGroup -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize $unlimited -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | ? { ($_.Alias -notmatch "DiscoverySearchMailbox") -and ($_.PrimarySmtpAddress -notmatch $regexdomain) } | Select Alias, DisplayName, @{n = "EmailAddresses"; e = { $_.EmailAddresses -join ";" } }, ExternalEmailAddress, FirstName, HiddenFromAddressListsEnabled, LastName, LegacyExchangeDn, Name, PrimarySmtpAddress, RecipientType, @{n = "Phone"; e = { "" } }, @{n = "MobilePhone"; e = { "" } }, @{n = "Company"; e = { "" } }, @{n = "Title"; e = { "" } }, @{n = "Department"; e = { "" } }, @{n = "Office"; e = { "" } }, guid
            }
            RemoteUserMailbox {
                Write-Host "    Exporting RemoteMailboxes" -ForegroundColor Yellow
                Set-AdServerSettings -ViewEntireForest $True
                $galTemp = Get-RemoteMailbox -resultsize $unlimited -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | ? { ($_.Alias -notmatch "DiscoverySearchMailbox") -and ($_.PrimarySmtpAddress -notmatch $regexdomain) } | Select Alias, DisplayName, @{name = "EmailAddresses"; e = { $_.EmailAddresses -join ";" } }, @{Name = "ExternalEmailAddress"; Expression = { $(($_.RemoteRoutingAddress).SmtpAddress) } }, FirstName, HiddenFromAddressListsEnabled, LastName, LegacyExchangeDn, Name, PrimarySmtpAddress, @{n = "RecipientType"; e = { $_.RecipientTypeDetails } }, @{n = "Phone"; e = { "" } }, @{n = "MobilePhone"; e = { "" } }, @{n = "Company"; e = { "" } }, @{n = "Title"; e = { "" } }, @{n = "Department"; e = { "" } }, @{n = "Office"; e = { "" } }, guid
                $galTempp += finaluserprop($galTemp)
            }

        }

    }

#empty company - fill with domain : FIXME
#    $galTempp | ForEach-Object {
#        if ($_.'Company' -eq "") {
#            $_-'PrimarySmtpAddress'.Split("@")[1] 
#            $_-'Company' = $_-'PrimarySmtpAddress'.Split("@")[1]
#        }
#    }    

    $galTempp | Export-Csv $outputcsv -Append -NoTypeInformation -Delimiter ","
    Write-Log -Message "$($galTempp.Count) Objects were exported,," -Level INFO -logfile $logfi

}

######################################
#  Creating the Mailcontacts in EXO  #
######################################

function import () {
    Write-Host "We need to authenticate to EXO to create the MailContacts" -ForegroundColor black -BackgroundColor Yellow
    auth_Online
       
    If ($RecipientTypes) {
        $RecipientTypeFilter = ($RecipientTypes | % { $_ }) -join "|"
        Write-Host "Applying RecipientTypes filter $($RecipientTypeFilter) to import list." -ForegroundColor Yellow
        $SourceGAL = Import-Csv $inputcsv | ? { $_.RecipientType -match $RecipientTypeFilter }
    }
    else {
        $SourceGAL = Import-Csv $inputcsv -Delimiter "," -Encoding UTF8
    }

    $SourceGAL = $SourceGAL | ? { ($_.PrimarySmtpAddress -ne "") }

    $objCount = $SourceGAL.Count
    $c = 0
    foreach ($gal in $SourceGAL) {
        $c++
        $currentTime = (Get-Date)
        if (($currentTime - $startTime).minutes -gt 15) {
            Start-Sleep -Seconds 500
            $startTime = (Get-Date)
        }
        Write-Progress -Activity "$c out of $objCount" -Status $gal.PrimarySmtpAddress -PercentComplete (($c / $objCount) * 100)
        $Alias = $gal.Alias
        $DisplayName = $gal.DisplayName
        if ($gal.LegacyExchangeDN) {
            $LegDN = "x500:" + $gal.LegacyExchangeDN
        }

        [array]$EmailAddresses = $gal.EmailAddresses.Split(";")
        [array]$EmailAddresses = $EmailAddresses -match '@'
        [array]$EmailAddresses = $EmailAddresses -notmatch 'SIP'
        if ($gal.LegacyExchangeDN) {
            [array]$EmailAddresses = $EmailAddresses + $LegDN
        }
        [array]$EmailAddresses = $EmailAddresses | Sort -Unique    
        $EmailAddressesCount = $EmailAddresses.Count
        $FirstName = $gal.FirstName
        $LastName = $gal.LastName
        $Name = $gal.Name
        $Phone = $gal.Phone
        $MobilePhone = $gal.MobilePhone
        $Company = $gal.Company
        $Title = $gal.Title
        $Department = $gal.Department
        $Office = $gal.Office
        # Needed to convert "True" or "False" value from text to Boolean
        $HiddenFromAddressListsEnabled = [System.Convert]::ToBoolean($gal.HiddenFromAddressListsEnabled)
        $checkrecip = (Get-Recipient -Identity $($gal.PrimarySmtpAddress) -ErrorAction SilentlyContinue).RecipientType
        if ($checkrecip -match 'MailUser|MailContact') {
            $ExternalEmailAddress = ($gal.ExternalEmailAddress).split(':')[1]
        }
        else {
            $ExternalEmailAddress = $gal.PrimarySmtpAddress
        }

        
        if (!$checkrecip) {
            try {
                $CreateMailContact = {
                    param($Alias,
                        $DisplayName,
                        $ExternalEmailAddress,
                        $FirstName,
                        $LastName,
                        $Name,
                        $HiddenFromAddressListsEnabled,
                        $EmailAddresses,
                        $Phone, 
                        $MobilePhone,
                        $Company,
                        $Title,
                        $Department,
                        $Office
                    )

                    New-MailContact -Alias $Alias -DisplayName $DisplayName -ExternalEmailAddress $ExternalEmailAddress -FirstName $FirstName -LastName $LastName -Name $Name -ea SilentlyContinue
                    
                }

                $SetMailContact = {
                    param(
                        $Alias,
                        $Name,
                        $HiddenFromAddressListsEnabled,
                        $EmailAddresses
                    )

                    Set-MailContact $Alias -HiddenFromAddressListsEnabled $HiddenFromAddressListsEnabled -EmailAddresses $EmailAddresses -Name $Name -ea SilentlyContinue
                    
                }
                
                Write-Host -NoNewline "Trying to create new contact for " -ForegroundColor Yellow; Write-Host -ForegroundColor Cyan $ExternalEmailAddress
                #Write-Host "   Updating $($EmailAddressesCount) proxy addresses..." -ForegroundColor Green
                Invoke-Command -Session (Get-PSSession) -ScriptBlock $CreateMailContact -ArgumentList $Alias, $DisplayName, $ExternalEmailAddress, $FirstName, $LastName, $Name, $HiddenFromAddressListsEnabled, $EmailAddresses | Out-Null
                #Start-Sleep -Milliseconds 200
                Invoke-Command -Session (Get-PSSession) -ScriptBlock $SetMailContact -ArgumentList $Alias, $Name, $HiddenFromAddressListsEnabled, $EmailAddresses | Out-Null
                Start-Sleep -Milliseconds 300
                Set-Contact $Alias -Phone $Phone -MobilePhone $MobilePhone -Title $Title -Department $Department -Office $Office -Company $Company -ea SilentlyContinue | Out-Null
                If ($SyncAttribute) {
                    $cmd = "Set-MailContact $Alias -$SyncAttribute $SyncAttributeValue"
                    Invoke-Expression $cmd
                }
                Write-Host "   The MailContact was created successfuly" -ForegroundColor Green
                Write-Log -Message "$ExternalEmailAddress MailboxContact was added, $ExternalEmailAddress,$DisplayName" -Level INFO -logfile $logfi
            }
            catch {
                Write-Host "We coundn't create the MailContact for $($gal.PrimarySmtpAddress)" -ForegroundColor Red
                Write-Log -Message "$ExternalEmailAddress couldn't be created, $ExternalEmailAddress,$DisplayName" -Level ERROR -logfile $logfi
            }
        }
        elseif ($checkrecip -eq 'MailContact') {
            try {
                Set-MailContact $Alias -HiddenFromAddressListsEnabled $HiddenFromAddressListsEnabled -EmailAddresses $EmailAddresses -Name $Name -ea SilentlyContinue | Out-Null
                Set-Contact $Alias -Phone $Phone -MobilePhone $MobilePhone -Title $Title -Department $Department -Office $Office -Company $Company -ea SilentlyContinue | Out-Null
                If ($SyncAttribute) {
                    $cmd = "Set-MailContact $Alias -$SyncAttribute $SyncAttributeValue"
                    Invoke-Expression $cmd
                    Write-Host "The MailContact $ExternalEmailAddress already exist and was updated successfuly" -ForegroundColor Green
                    Write-Log -Message "$ExternalEmailAddress UserMaibox was updated, $ExternalEmailAddress,$DisplayName" -Level INFO -logfile $logfi
                }
            }
            catch {
                Write-Host "We tried to update the MailContact $ExternalEmailAddress unsuccessfuly :-( " -ForegroundColor Red
                Write-Log -Message "We tried to update the MailContact $ExternalEmailAddress unsuccessfuly :-(, $ExternalEmailAddress,$DisplayName" -Level ERROR -logfile $logfi
            }
        }
        else {
            Write-Host "The address $($gal.PrimarySmtpAddress) already exist as $checkrecip and cannot be updated" -ForegroundColor Red
            Write-Log -Message "$ExternalEmailAddress UserMaibox already exist as $checkrecip, $ExternalEmailAddress,$DisplayName" -Level INFO -logfile $logfi
        }
    }
} 


############################################################
#  Removing the MailContacts based on the Attribute value  #
############################################################

function deleteMailContacts () {
    Write-Host "We need to authenticate to EXO to delete the MailContacs with the value $SyncAttributeValue in $SyncAttribute " -ForegroundColor black -BackgroundColor Yellow
    auth_Online
    $recipwithAtt = Get-MailContact | ? { $_.CustomAttribute10 -eq "Sync" }
    $d = 0
    if ($recipwithAtt) {
        foreach ($rec in $recipwithAtt) {
            $d++
            Write-Progress -Activity "$d out of $($recipwithAtt.count)" -Status $d -PercentComplete (($d / $recipwithAtt.count) * 100)
            try {
                Write-Host "Deleting the Mailcontact $($rec.PrimarySmtpAddress)" -ForegroundColor Yellow
                Remove-MailContact -Identity $rec.identity -Confirm:$false
                Write-Log -Message "Deleting the Mailcontact $($rec.PrimarySmtpAddress), $($rec.PrimarySmtpAddress),$($rec.DisplayName)" -Level INFO -logfile $logfi
            }
            catch {
                Write-Host "We couldn't delete the Mailcontact $($rec.PrimarySmtpAddress)" -ForegroundColor Red
                Write-Log -Message "We couldn't delete the Mailcontact $($rec.PrimarySmtpAddress), $($rec.PrimarySmtpAddress),$($rec.DisplayName)" -Level ERROR -logfile $logfi
            }
        }
    }
    else {
        Write-Host "No Mailcontact with the value $SyncAttributeValue in the $SyncAttribute were found" -ForegroundColor Red
        Write-Log -Message "No Mailcontact with the value $SyncAttributeValue in the $SyncAttribute were found, ," -Level WARN -logfile $logfi
    }

}

##############################
#  Delete disabled accounts  #
##############################

function disables() {
    Write-Host "We are checking your disable accounts to delete the EXO MailContacs associated with them" -ForegroundColor Yellow -BackgroundColor Red
    Write-Host "   Authenticating to On premises to get the disable accounts`n" -ForegroundColor Yellow
    auth_OnPremises
    $dis_Accounts = @()
    $dis_Accounts = Get-Mailbox -ResultSize Unlimited | ? { ($_.OrganizationalUnit -match 'HKEnterprise.*Disabled-Accounts') } | select DisplayName, PrimarySmtpAddress
    Write-Host "   Authenticating to Exchange Online to delete the MailContacts with disabled accounts`n" -ForegroundColor Yellow
    auth_Online
    if ($dis_Accounts) {
        foreach ($dis in $dis_Accounts) {
            try {
                $checkMC = Get-MailContact -Identity $dis.PrimarySmtpAddress -ErrorAction SilentlyContinue  | ? { $_.CustomAttribute10 -eq 'Sync' }
                if ($checkMC) {
                    Write-Host "   The MailContact $($dis.PrimarySmtpAddress) exist in your EXO tenant and going to be delete because is a disable account on premises" -ForegroundColor Red
                    Write-Log -Message "MailContact deleted - Disabled Account on premises, $($dis.PrimarySmtpAddress), $($dis.DisplayName) " -Level WARN -logfile $logfi
                    Remove-MailContact -Identity $dis.PrimarySmtpAddress -Confirm:$false
                } 
            }
            catch {
                Write-Host "   We faced some issues running the Get-MailContact -Identity $($dis.PrimarySmtpAddress) command " -ForegroundColor Red -BackgroundColor Yellow
                Write-Log -Message "We faced some issues running the Get-MailContact -Identity $($dis.PrimarySmtpAddress) command " -Level ERROR -logfile $logfi
            }
        }
    }

}


[System.GC]::Collect()

switch ($operation) {
    "export" {
        get_Excluded_Domains
        export
    }
    "import" {
        $script:startTime = (Get-Date)
        import
        # disables
    }
    "delete" {
        if ($ConfirmDeletes) {
            deleteMailContacts
            #disables
        }
        else {
            Write-Host "If you want to delete the MailContacts then select the parameter ConfirmDeteles" -ForegroundColor Red
        }
            
    }

}



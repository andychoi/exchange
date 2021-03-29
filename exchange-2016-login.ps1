$Session2016 = ""
$Session2016 = (Get-PSSession).Name
if ($Session2016 -ne $null) {
    Remove-PSSession -Session (Get-PSSession) -ErrorAction SilentlyContinue
}

$$Credentials2016 = Get-Credential
$Session2016 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://owa.hisna.com/PowerShell/ -Authentication Kerberos -Credential $$Credentials2016
#$Session2016 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://irvhaeexmp03v/PowerShell/ -Authentication Kerberos -Credential 

Import-PSSession $Session2016

#Import-Module ActiveDirectory


#Remove-PSSession $Session2016


$domainadmins = Get-ADGroupMember "domain admins"
$enterpriseadmins = Get-ADGroupMember "enterprise admins"

$DOMAIN_DATA = @()
$remain_DATA = @()

$skip_admin = $true

$routingdomain = "mercoline.mail.onmicrosoft.com"

$Domain_before = "mercoline.de"

$Domain_after = "mercoline-edi.de"

#$OUfilter = ""

$allMbx = Get-ADUser -LDAPFilter "(msExchMailboxGuid=*)" -Properties UserPrincipalName,DistinguishedName,ObjectGuid,mail,mailnickname,proxyaddresses,msExchMailboxGuid

$mbxs = $allMbx | where { $_.Samaccountname -notin $domainadmins.Samaccountname -and $_.Samaccountname -notin $enterpriseadmins.Samaccountname }


foreach ($mbx in $mbxs) {

 $Local_UPN = "" ; 

$Local_SMTP = "" ;

    $Samaccountname = $mbx.SamAccountName
      $mailnickname = $mbx.mailnickname
  [guid]$ObjectGUID = $mbx.ObjectGUID
      $ExchangeGUID = [Guid]$mbx.msExchMailboxGuid

$primarysmtpaddress = $mbx.proxyaddresses | where { $_ -cmatch "^SMTP:" }
          $new_smtp = $primarysmtpaddress -replace "mercoline.de","mercoline-edi.de"
          $old_smtp = $primarysmtpaddress -replace "mercoline-edi.de","mercoline.de"
             $Wrong = $new_smtp -creplace "SMTP:","smtp:"
         $Wrong_sip = $old_smtp -creplace "SMTP:","SIP:"
           $new_sip = $new_smtp -creplace "SMTP:","SIP:"

 Set-ADUser -identity $Samaccountname -Remove @{ProxyAddresses="$Wrong_sip"} -ErrorAction stop -Confirm:$false

 Set-ADUser -identity $Samaccountname -Add @{ProxyAddresses="$new_sip"} -ErrorAction stop -Confirm:$false

}

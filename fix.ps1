foreach ($mbx in $mbxs) {

 $Local_UPN = "" ; 

$Local_SMTP = "" ;

    $Samaccountname = $mbx.SamAccountName
      $mailnickname = $mbx.mailnickname
  [guid]$ObjectGUID = $mbx.ObjectGUID
      $ExchangeGUID = [Guid]$mbx.msExchMailboxGuid

$primarysmtpaddress = $mbx.proxyaddresses | where { $_ -cmatch "^SMTP:" }
          $new_smtp = $primarysmtpaddress -replace "mercoline.de","mercoline-edi.de"
             $Wrong = $new_smtp -creplace "SMTP:","smtp:"

 Set-ADUser -identity $Samaccountname -Remove @{ProxyAddresses="$Wrong"} -ErrorAction stop -Confirm:$false

 Set-ADUser -identity $Samaccountname -Add @{ProxyAddresses="$new_smtp"} -ErrorAction stop -Confirm:$false

}

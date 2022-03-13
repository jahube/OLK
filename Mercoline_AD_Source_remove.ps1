Import-Module Activedirectory

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
# $allMbx = Get-ADUser -LDAPFilter "(msExchMailboxGuid=*)" -Search $OUfilter -Properties UserPrincipalName,DistinguishedName,ObjectGuid,mail,mailnickname,proxyaddresses,msExchMailboxGuid

$mbxs = $allMbx | where { $_.Samaccountname -notin $domainadmins.Samaccountname -and $_.Samaccountname -notin $enterpriseadmins.Samaccountname }

$mbxs.Count

$path = "c:\temp"

foreach ($mbx in $mbxs) {

 $Local_UPN = "" ; 

$Local_SMTP = "" ;

    $Samaccountname = $mbx.SamAccountName
      $mailnickname = $mbx.mailnickname
  [guid]$ObjectGUID = $mbx.ObjectGUID
      $ExchangeGUID = [Guid]$mbx.msExchMailboxGuid
$primarysmtpaddress = $mbx.proxyaddresses | where { $_ -cmatch "^SMTP:" }
               $SIP = $mbx.proxyaddresses | where { $_ -match "^SIP:" }
         $Local_UPN = $mbx.userprincipalname | where { $_ -match "$Domain_before$" }
        $Local_SMTP = $primarysmtpaddress | where { $_ -match "$Domain_before$"  }
            $PREFIX = ($primarysmtpaddress -split '@')[0]
               $UPN = $mbx.userprincipalname

        $item = New-Object -TypeName PSObject       
        $item | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $userprincipalname
        $item | Add-Member -MemberType NoteProperty -Name primarysmtpaddress -Value $primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name mailnickname -Value $mailnickname
        $item | Add-Member -MemberType NoteProperty -Name Samaccountname -Value $Samaccountname
        $item | Add-Member -MemberType NoteProperty -Name PREFIX -Value $PREFIX
        $item | Add-Member -MemberType NoteProperty -Name distinguishedname -Value $mbx.distinguishedname
        $item | Add-Member -MemberType NoteProperty -Name ObjectGUID -Value $ObjectGUID
        $item | Add-Member -MemberType NoteProperty -Name ExchangeGUID -Value $ExchangeGUID


IF (!($Local_UPN) -and !($Local_SMTP))

{
        Write-host "SMTP: $primarysmtpaddress | UPN: $UPN - Korrekt`n`n" -ForegroundColor Green

        $item | Add-Member -MemberType NoteProperty -Name old_SMTP -Value $primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name new_SMTP -Value $primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name result_SMTP -Value "SMTP korrekt"

        $item | Add-Member -MemberType NoteProperty -Name old_UPN -Value $UPN
        $item | Add-Member -MemberType NoteProperty -Name new_UPN -Value $UPN
        $item | Add-Member -MemberType NoteProperty -Name result_UPN -Value "UPN korrekt"

        $DOMAINDATA += $item ;
}


<#
Set-ADUser USER -Add @{ProxyAddresses="SMTP:neue_primaryaddress@mercoline.de"}

Set-ADUser USER -Remove @{ProxyAddresses="SMTP:old_primaryaddress@datagroup.de"}

Set-ADUser USER -Add @{ProxyAddresses="smtp:old_primaryaddress@datagroup.de"}

Set-ADUser USER -Replace @{mail="neue_primaryaddress@mercoline.de"}
#>

 # primarysmtpaddress

 IF ($Local_SMTP)

 {
        $SMTP_After = $Local_SMTP -Replace ( $Domain_before, $Domain_after )

        $Mail = $SMTP_After -Replace "smtp:"

         Write-host "SMTP: before $primarysmtpaddress | " -ForegroundColor cyan -NoNewline

        Try {

 Set-ADUser -identity $Samaccountname -Add @{ProxyAddresses="$SMTP_After"} -ErrorAction stop -Confirm:$false

 Set-ADUser -identity $Samaccountname -Replace @{mail="$Mail"} -ErrorAction stop -Confirm:$false

 Set-ADUser -identity $Samaccountname -Remove @{ProxyAddresses="$Local_SMTP"} -ErrorAction stop -Confirm:$false

 $proxys_left =  Get-ADUser -identity $Samaccountname -Properties ProxyAddresses | select -ExpandProperty ProxyAddresses | where { $_ -match $Domain_before }

 foreach($remain in $proxys_left) {
 
 Set-ADUser -identity $Samaccountname -Remove @{ProxyAddresses="$remain"} -ErrorAction stop -Confirm:$false

 $remain_DATA += $remain ;

 }
 
 Write-host "SMTP: after $SMTP_After - SUCCESS`n" -ForegroundColor green

        $item | Add-Member -MemberType NoteProperty -Name old_SMTP -Value $primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name new_SMTP -Value $SMTP_After
        $item | Add-Member -MemberType NoteProperty -Name result_SMTP -Value "success"

            }

          catch

            {

 Write-host "SMTP: after $primarysmtpaddress - SMTP $SMTP_After UPDATE FAILED" -ForegroundColor red -BackgroundColor Yellow

        $item | Add-Member -MemberType NoteProperty -Name old_SMTP -Value $primarysmtpaddress
        $item | Add-Member -MemberType NoteProperty -Name new_SMTP -Value $SMTP_After
        $item | Add-Member -MemberType NoteProperty -Name result_SMTP -Value $Error[0].Exception.Message

            }

}

 # userprincipalname

 IF ($Local_UPN)

 {
        Write-host "UPN: before $UPN | " -ForegroundColor cyan -NoNewline

        $UPN_After = $Local_UPN -replace ( $Domain_before, $Domain_after )

        Try {

 Set-ADUser -identity $Samaccountname -Replace @{userprincipalname="$UPN_After"} -ErrorAction stop -Confirm:$false

 Write-host "UPN: after change $UPN_After - SUCCESS`n`n" -ForegroundColor green

        # userprincipalname

        $item | Add-Member -MemberType NoteProperty -Name old_UPN -Value $UPN
        $item | Add-Member -MemberType NoteProperty -Name new_UPN -Value $UPN_After
        $item | Add-Member -MemberType NoteProperty -Name result_UPN -Value "UPN korrekt"

            }

          catch

            {
            
 Write-host "UPN: after $UPN - UPN $UPN_After UPDATE FAILED`n`n" -ForegroundColor red -BackgroundColor Yellow

        $item | Add-Member -MemberType NoteProperty -Name old_UPN -Value $UPN
        $item | Add-Member -MemberType NoteProperty -Name new_UPN -Value $UPN_After
        $item | Add-Member -MemberType NoteProperty -Name result_UPN -Value $Error[0].Exception.Message

            }

}

# SIP

$SIP = $mbx.proxyaddresses | where { $_ -match "^SIP:" }


IF ($SIP)

 {
        Write-host "SIP: before $SIP | " -ForegroundColor cyan -NoNewline

        $SIP_After = $SIP -replace ( $Domain_before, $Domain_after )

        Try {

 Set-ADUser -identity $Samaccountname -Remove @{ProxyAddresses="$SIP"} -ErrorAction stop -Confirm:$false

 Set-ADUser -identity $Samaccountname -Add @{ProxyAddresses="$SIP_After"} -ErrorAction stop -Confirm:$false

 Write-host "SIP: after change $SIP_After - SUCCESS`n`n" -ForegroundColor green

        # SIP

        $item | Add-Member -MemberType NoteProperty -Name old_SIP -Value $SIP
        $item | Add-Member -MemberType NoteProperty -Name new_SIP -Value $SIP_After
        $item | Add-Member -MemberType NoteProperty -Name result_SIP -Value "SIP korrekt"

            }

          catch

            {
            
 Write-host "SIP: after $SIP - SIP $SIP_After UPDATE FAILED`n`n" -ForegroundColor red -BackgroundColor Yellow

        $item | Add-Member -MemberType NoteProperty -Name old_SIP -Value $SIP
        $item | Add-Member -MemberType NoteProperty -Name new_SIP -Value $SIP_After
        $item | Add-Member -MemberType NoteProperty -Name result_SIP -Value $Error[0].Exception.Message

            }

}

      $DOMAIN_DATA += $item ;

}

# Export CSV Data

if (!(Test-Path $path)) { mkdir $path }

$datestamp = Get-Date -Format yyyy.MM.dd_HH.MM
$filepath = $path + '\AD_DOMAIN_Update_DATA_' + $datestamp + '.CSV'
$DOMAIN_DATA | Export-Csv -Path $filepath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force

#$DOMAIN_DATA | Export-Csv -Path "$path\AD_DOMAIN_Update_DATA_$datestamp.CSV" -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force

$remain_DATA | Export-Csv -Path "$path\remain_DATA_$datestamp.CSV" -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force

# END
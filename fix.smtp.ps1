$mbxs = get-remotemailbox -resultsize unlimited

$array = $mbxs | where { $_.primarysmtpaddress -match "mercoline.de" }

foreach ($user in $array) {

$new_smtp = $user.primarysmtpaddress -replace "mercoline.de","mercoline-edi.de"
$old_smtp = $user.primarysmtpaddress

set-remotemailbox $user.userprincipalname -primarysmtpaddress "$new_smtp"

set-remotemailbox $user.userprincipalname -proxyaddresses @{remove="$old_smtp"} 

}
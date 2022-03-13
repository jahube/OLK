
$Domain_before = "mercoline.de"

$Domain_after = "mercoline-edi.de"

$groups = get-distributiongroup -ResultSize unlimited -organizationalunit "OU=_Mercoline_Verteilerlisten,DC=mercoline,DC=local"

$groups |ft primarysmtpaddress,EmailAddresses,windowsemailaddress

foreach ($group in $groups) {

Set-DistributionGroup $group.distinguishedname -EmailAddressPolicyEnabled:$false

$SMTP =$group.PrimarySmtpAddress

$NEW_SMTP = $SMTP -replace ( $Domain_before, $Domain_after )

Set-DistributionGroup $group.distinguishedname -PrimarySmtpAddress $NEW_SMTP -Confirm:$false -force

Set-DistributionGroup $group.distinguishedname -WindowsEmailAddress $NEW_SMTP -Confirm:$false -force

}
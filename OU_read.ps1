$OUs = Get-ADOrganizationalUnit -filter * -SearchBase "OU=groups,OU=primeo-energie,DC=pe,DC=ch" -Properties  managedby, businessCategory | Select-Object distinguishedname, name, managedby, businessCategory, ObjectGUID
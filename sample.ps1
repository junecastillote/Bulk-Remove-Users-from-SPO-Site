# Connect to the SPO Admin Site
Connect-SPOService -Url 'https://<tenant>-admin.sharepoint.com'

# Import the list of users to remove.
$usersToDelete = Import-Csv .\UsersToRemove.csv

## Sample CSV File
# "Url","LoginName"
# "https://contoso.sharepoint.com/sites/SITENAME001","user_001@domain.com"
# "https://contoso.sharepoint.com/sites/SITENAME002","user_002@domain.com"

# Run the script
.\Remove-SPOSiteUsersBulk.ps1 -InputObject $usersToDelete -OutputDirectory .\report -Live -Confirm:$false

# If you want to exclude one or more users:
$excludeUser = @('service_account@domain.com','admin_account@domain.com')
# Run the script
.\Remove-SPOSiteUsersBulk.ps1 -InputObject $usersToDelete -OutputDirectory .\report -Live -Confirm:$false -ExcludeUser $excludeUser

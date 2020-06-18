# Set mail on general tab

# Data received from caller
Param (
    [Parameter(Mandatory,ValueFromPipeline)]
    [String[]]$users,
    [Parameter(Mandatory)]
    [String]$outputFile
)

# Loop through users setting mail. Add user to outputFile if errors and reason. 
ForEach ($user in $users) 
{
    # Avoid empty entries
    if ($user -ne "") {
        #Set the Email in General tab from the UPN
        Get-ADUser $user -Properties mail | % {Set-ADUser $_ -EmailAddress ($_.UserPrincipalName)}
    }
}

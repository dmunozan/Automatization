# Export details for NH mail

# Data received from caller
Param (
    [Parameter(Mandatory,ValueFromPipeline)]
    [String[]]$users,
    [Parameter(Mandatory)]
    [String]$parentPath,
    [Parameter(Mandatory)]
    [String]$outputFile
)

$password = Get-Content -Path "$($parentPath)\Files\password.txt"

#Get data for NH mail and save it in Output
$Output = ForEach ($user in $users)
{
    $currentUser = Get-ADUser -Filter 'SamAccountName -like $user' -Properties DisplayName, mailNickname, UserPrincipalName, SamAccountName

    [pscustomobject] @{
                Name = $currentUser.DisplayName
                'E-mail' = $currentUser.mailNickname
                '@Domain' = $currentUser.UserPrincipalName.Substring($currentUser.UserPrincipalName.IndexOf('@'))
                User = $currentUser.SamAccountName
                Password = $password
            }
}

#Export output to CSV file
$Output | Export-CSV "$($parentPath)\MailData\NH_Mail-$($env:username).csv" -NoTypeInformation

# Get username from Display Name

# Path to the script, import file and delete contents on FoundUserNames file
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$parentPath = (Get-Item $scriptPath).Parent.FullName
$users = Import-Csv -Path "$($parentPath)\Files\ExpectedDisplayNames-$($env:username).csv"
"" | Out-File "$($parentPath)\Files\FoundUsernames-$($env:username).csv" -NoNewline

# Loop through users fetching userName from AD using ExpectedDisplayNames and adding it to FoundUserNames file.
# Add expected DisplayName if not found
$users|Foreach{
    # Avoid empty entries
    if ($_.DisplayName -ne "") {
        # Look for userName using expected DisplayName
        $user = Get-ADUser -Filter "DisplayName -like '$($_.DisplayName)*'" -Properties SamAccountName
        
        # Get userName or expected DisplayName if userName not found
        $data = ""
        if ($user -eq $null) {
            $data = $_.DisplayName
        } else {
            $data = $user.SamAccountName
        }

        # Append data to FoundUserNames file
        $data | Out-File "$($parentPath)\Files\FoundUsernames-$($env:username).csv" -Append
    }
}

# Set default password and require change at next logon

# Data received from caller
Param (
    [Parameter(Mandatory,ValueFromPipeline)]
    [String[]]$users,
    [Parameter(Mandatory)]
    [String]$parentPath,
    [Parameter(Mandatory)]
    [String]$outputFile
)

# Set the default password
$stringPassword = Get-Content -Path "$($parentPath)\Files\password.txt"
$password = ConvertTo-SecureString -AsPlainText $stringPassword -Force 

# Loop through users setting the password. Add user to outputFile if errors and reason.
ForEach ($user in $users) {
    # Avoid empty entries
    if ($user -ne "") {
        # Set the default password for the current user
        Set-ADAccountPassword -Identity $user -NewPassword $password -Reset
    
        # Set the change password option to true
        Set-AdUser -Identity $user -ChangePasswordAtLogon $true
    }
}

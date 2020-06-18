# Move users to $ou

# Data received from caller
Param (
    [Parameter(Mandatory,ValueFromPipeline)]
    [String[]]$users,
    [Parameter(Mandatory)]
    [String]$ou,
    [Parameter(Mandatory)]
    [String]$outputFile
)

# Loop through users moving them to $ou. Add user to outputFile if errors and reason. 
ForEach ($user in $users) 
{
    # Avoid empty entries
    if ($user -ne "") {
        # Move object to $ou
        Get-ADUser $user | Move-ADObject -TargetPath $ou
    }
}

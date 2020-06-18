# Config AD Users using UserData file

# Path to the script, import file and delete contents on result file
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$parentPath = (Get-Item $scriptPath).Parent.FullName
$data = Import-Csv -Path "$($parentPath)\Files\UserData-$($env:username).csv"
$users = $data.Username

# Saves the OU that is located on the first position (if more than one user $ou becomen an array throwing an error on moveObject paramaters)
$ou = $data.OU[0]
$outputFile = "$($parentPath)\Files\Result-$($env:username).csv"
"" | Out-File $outputFile -NoNewline


."$($parentPath)\Scripts\setPassword.ps1" -users $users -parentPath $parentPath -outputFile $outputFile
."$($parentPath)\Scripts\addMail.ps1" -users $users -outputFile $outputFile
."$($parentPath)\Scripts\moveObject.ps1" -users $users -ou $ou -outputFile $outputFile
."$($parentPath)\Scripts\exportDataForNHMail.ps1" -users $users -parentPath $parentPath -outputFile $outputFile

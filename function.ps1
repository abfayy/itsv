<#
 .Name: Functions.ps1
 .Author: Abdullah Fayyaz
 .Deployed to: ---
 .Summary: This loads function for all MultiGeo SPO Sessions creation and more
 .Tracked: <Git Repository> 
#>


Write-Host "Loading functions....."
$sec = Get-Content .\secrets.txt
#Connecting to NAM 
function CreateSPONAMSession
{
    #Write-Host -foreground Yellow "You are not connected!"
    $username = "<##Input Service Account>"
    $password = $sec #"$Env:secret" #AF: ADD Catch  
    #$password
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $(convertto-securestring $Password -asplaintext -force)
    try
    {
        Connect-SPOService -Url "<##Tenant Url for first multi-geo>" -Credential $cred
        Write-Host "Connect-SPONAMService is successful!"
    }catch{
        Write-Host "Error in connection"
        Send-MailMessage -SmtpServer "<##Input: SMPT server Url | don't forget to whitelist the server on SMTP relay server>" -From "Guru@spacesharepoint.com" -Subject "Connection error with SPOnline NAM Service" -Body "Hi, `nThe Connect Script ran on PWEBQA01 failed.`n$($_.Exception.Message)`n$($_.Exception.ItemName)`nPlease check again.`n - Automation Guru"
        $ErrorMessage = $_.Exception.Message
        $ErrorMessage | out-file .\error.log -append
        $FailedItem = $_.Exception.ItemName
        $FailedItem  | out-file .\error.log -append
    }finally{

        "$(Get-Date) This script SPONAMService made a read attempt" | out-file c:\scripts\error.log -append

    }
}
#Run using 
#Create-SPOSession


#Connecting to CAN
function CreateSPOCANSession
{
    #Write-Host -foreground Yellow "You are not connected!"
    $username = "<##Input Service Account>"
    $password = $sec #"$Env:secret" #AF: ADD Catch 
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $(convertto-securestring $Password -asplaintext -force)
    try
    {
        Connect-SPOService -Url "<##Tenant Url for second multi-geo>" -Credential $cred
        Write-Host "Connect-SPOCANService is successful!"
    }catch{
        Write-Host "Error in connection"
        Send-MailMessage -SmtpServer "<##Input: SMPT server Url | don't forget to whitelist the server on SMTP relay server>" -From "Guru@spacesharepoint.com" -Subject "Connection error with CAN SPOnline Service" -Body "Hi, `nThe Connect Script ran on PWEBQA01 failed.`n$($_.Exception.Message)`n$($_.Exception.ItemName)`nPlease check again.`n - Automation Guru"
        $ErrorMessage = $_.Exception.Message
        $ErrorMessage | out-file .\error.log -append
        $FailedItem = $_.Exception.ItemName
        $FailedItem  | out-file .\error.log -append
    }finally{
        "$(Get-Date) This script SPOCANService made a read attempt" | out-file c:\scripts\error.log -append

    }
    
}

#Connecting to MEA Tenant
function CreateSPOMEASession
{
    #Write-Host -foreground Yellow "You are not connected!"
    $username = "<##Input Service Account>"
    $password = $sec #"$Env:secret" #AF: ADD Catch 
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $(convertto-securestring $Password -asplaintext -force)
    try
    {
        Connect-SPOService -Url "<##Tenant Url for third multi-geo>" -Credential $cred
        Write-Host "Connect-SPOMEAService is successful!"
    }catch{
        Write-Host "Error in connection"
        Send-MailMessage -SmtpServer "<##Input: SMPT server Url | don't forget to whitelist the server on SMTP relay server>" -From "Guru@spacesharepoint.com" -Subject "Connection error with SPOnline MEA Service" -Body "Hi, `nThe Connect Script ran on PWEBQA01 failed.`n$($_.Exception.Message)`n$($_.Exception.ItemName)`nPlease check again.`n - Automation Guru"
        $ErrorMessage = $_.Exception.Message
        $ErrorMessage | out-file .\error.log -append
        $FailedItem = $_.Exception.ItemName
        $FailedItem  | out-file .\error.log -append
    }finally{
        "$(Get-Date) This script SPOMEAService made a read attempt" | out-file .\error.log -append

    }
}
        "$(Get-Date) This script Completely connected" | out-file .\error.log -append
        "Function was loaded successfully" | out-file .\error.log -append
        
Write-Host "Functions Loaded!!!"

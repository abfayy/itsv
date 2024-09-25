<#
 .Author: Abdullah Fayyaz
 .Comments: This script is deployed a server with a scheduled task running to allow legal to get access to a users onedrive. Search for ##Input to fill in values
 .Deployed to: <Deployment location
 .Tracked: <VSS source location>
#>

$sec = Get-Content .\secrets.txt (here the password was stored for the user account)

. .\Functions.ps1

function Get-ODRequest
{
    
    #<#002 - This portion retrives The Email address from the SharePoint List
    #Import the required DLL
    Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll'
    Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
    #OR
    Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll'

    #Mysite URL
    $site = '<##input: siteurl that contains the form>'
    
    #This was added to allow TLS 1.2 only
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    #Admin User Principal Name
    $admin = '<##input: username of service account>'
    $Password = $sec #"$Env:secret" #AF: ADD Catch 
    $spassword = ConvertTo-SecureString $Password –asplaintext –force 

    #Get the Client Context and Bind the Site Collection
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($site)

    #Authenticate
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin , $spassword)
    $context.Credentials = $credentials

    $list = $context.Web.Lists.GetByTitle('OneDriveRequest')
    $context.Load($list)

    $context.ExecuteQuery()

    #<#working
    foreach($lst in $list){Write-Host $lst.Title}

    #$caml="<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>achint</Value></Eq></Where></Query></View>"
    $caml="<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='test'>Requested</Value></Eq></Where></Query></View>"

    $cquery = New-Object Microsoft.SharePoint.Client.CamlQuery

    $cquery.ViewXml = $caml

    $listItems = $list.GetItems($cquery)

    $context.load($listItems)
    $error.Clear()
    #$error
    $context.executeQuery()
    if($error.Count -gt 0) { 
            "$(Get-date) The SharePoint list was not connected" | out-file .\error.log -append
            $error | out-file .\error.log -append 
            Exit
    }

    If($listItems.count -eq 0){
        "$(Get-date) This script made a read attempt" | out-file .\error.log -append
        "$(Get-date) Script exited; no items found" | out-file .\error.log -append
        Exit

    } else {
        Get-PSSession | Remove-PSSession

        New-PSSession -Name "NAM" #for US tenant
        New-PSSession -Name "CAN" #for Canada Tenant
        New-PSSession -Name "MEA" #for MEA tenant

        Invoke-Command -Session (Get-PSSession -Name "NAM") -ScriptBlock {
            $sec = Get-Content .\secrets.txt
            $sec
            . .\Functions.ps1
            CreateSPONAMSession
            $sps = Get-SPOSite -IncludePersonalSite $true -limit all  | Where-Object {$_.Url -like "*-my*"}
            $sps.count
        }

        Invoke-Command -Session (Get-PSSession -Name "CAN") -ScriptBlock {
            $sec = Get-Content .\secrets.txt
            $sec
            . .\Functions.ps1
            CreateSPOCANSession
            $sps = Get-SPOSite -IncludePersonalSite $true -limit all  | Where-Object {$_.Url -like "*-my*"}
            $sps.count
        }

        Invoke-Command -Session (Get-PSSession -Name "MEA") -ScriptBlock {
            $sec = Get-Content .\secrets.txt
            $sec
            . .\Functions.ps1
            CreateSPOMEASession
            $sps = Get-SPOSite -IncludePersonalSite $true -limit all  | Where-Object {$_.Url -like "*-my*"}
            $sps.count
        }


        foreach($listItem in $listItems)
        {
	        Write-Host "ID - " $listItem["ID"] "Title - " $listItem["Title"] "Status - " $listItem["Status"] 

            Write-Host $listItem["Title"]
            $email = $listItem["Title"].ToString()
            Write-Host $email


            <#00#>
            $scriptblock = {

                param([string] $email)

                #This is the email address string for the email message
                #
                $email2 = "<##Input: Email of users that need the email>
                $emailpre = $email.SubString(0, $email.IndexOf("@"))
                $urlstring = $emailpre.Replace(".","_")
                $urlstring
                #Adding permissions to the OneDrive
                $spu = $sps | Where-Object{$_.Url -like "*$($urlstring)_<##Input: domain of user email>*"}  #test.com as test_com
                $spu
                if ($null -ne $spu){
                    Set-SPOUser -Site $spu -LoginName "<##Input: Service account UPN>" -IsSiteCollectionAdmin $true
                    Set-SPOUser -Site $spu -LoginName "<##Input: UPN of account that needs access>" -IsSiteCollectionAdmin $true
                    If($spu.StorageUsageCurrent -gt "1"){
                        Send-MailMessage -SmtpServer "<##Input: SMPT server Url | don't forget to whitelist the server on SMTP relay server>" -From "Guru@spacesharepoint.com" -To $email2 -Subject "[OneDrive Data Check] Data is found for $($email)" -Body "Hi, `nThere is content found in the OneDrive for $($email). `nTotal Size of folder $($spu.StorageUsageCurrent).`nAccess has been updated for URL: $($spu.url).`n - Automation Guru"
                    } else 
                    {
                        Send-MailMessage -SmtpServer <##Input: SMPT server Url | don't forget to whitelist the server on SMTP relay server>" -From "Guru@spacesharepoint.com" -To $email2 -Subject "[OneDrive Data Check] Data not found for $($email)" -Body "Hi, `nNo content found in the OneDrive for $($email). `nTotal Size of folder $($spu.StorageUsageCurrent).`nAccess has been updated for URL: $($spu.url).`n - Automation Guru"
                        "$(Get-date) [OneDrive Data Check] Data not found for $($email)" | out-file .\error.log -append
                    }
                }else{
                    Send-MailMessage -SmtpServer <##Input: SMPT server Url | don't forget to whitelist the server on SMTP relay server>" -From "Guru@spacesharepoint.com" -To $email2 -Subject "[OneDrive Data Check] Data not found for $($email)" -Body "Hi, `nNo site found for $($email)`n - Automation Guru"
                }
            }        

            #On AD the user location was put in extensionAttribute15
	    $user = Get-ADUser -Filter 'UserPrincipalName -like $email' -Properties extensionAttribute15
            $pdl = $user.extensionAttribute15

            $error.Clear()

            switch ($pdl)
            {
                $null {
                    "This is US";
                    Invoke-Command -Session (Get-PSSession -Name "NAM") -ScriptBlock $scriptblock -ArgumentList $email
                    Break;
                }
                "ARE" {
                    "This is ARE";
                    Invoke-Command -Session (Get-PSSession -Name "MEA") -ScriptBlock $scriptblock -ArgumentList $email
                    Break
                }
                "CAN" {
                    "THis is CAN";
                    Invoke-Command -Session (Get-PSSession -Name "CAN") -ScriptBlock $scriptblock -ArgumentList $email
                    Break; 
                }
            }
            if($error.Count -gt 0) { 
                "$(Get-date) This script errored" | out-file .\error.log -append
                $error | out-file .\error.log -append 
            }else{
                #Updating Item in SharePoint
                #003#<#this is working
                $listItem["Status"] = "Completed"
                $listItem.Update()
                $context.ExecuteQuery()
                "$(Get-date) This script $($listItem["Title"].ToString()) completed" | out-file .\error.log -append
                #>#003#
            }
        }

    }

    

    Get-PSSession | Remove-PSSession
}

Get-ODRequest

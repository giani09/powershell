#---- Secure ClientSecret to .txt File ----##
#"client_secret in plaintext" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "C:\temp\password.txt"

#---- Varaibles ----#
$OutputPath = "C:\temp" # Set Path for Export
$lastdays = 30 # Set TimeStamp
$MailAddress = 'xxx' # Set the Mail Address
$tenantID = 'xxx' # Set the Tennant ID of the App Registration
$ClientID = 'xxx' # Set the Client ID of the App Registration
$scope = "https://graph.microsoft.com/.default"
$authority = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token"
$url = "https://graph.microsoft.com/v1.0/users/$MailAddress/messages" # URL for Object Messages in Azure Graph
#$ClientSecret = "xxx" # Set the ClientSecret (only for DEV!)
$ClientSecret = Get-Content "C:\temp\password3.txt" | ConvertTo-SecureString # Get secure ClientSecret

#---- Convert secure ClientSecret to PlainText ----#
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

#----- main-function -----#
function main() {
    $token = Authentication
    #$token
    getEmailAttachment $token
}

function Authentication {
    # Auth call
    $ReqTokenBody = @{
        Grant_Type    = "client_credentials"; # for Auth with ClientSecret
        client_id     = $ClientID;
        client_secret = $UnsecurePassword;
        Scope         = $scope;
    }
    return Invoke-RestMethod -Method POST -uri $authority -Body $ReqTokenBody
}


#---- for Test / DEV to Get AccessToken----#
<#$result = Invoke-RestMethod -Method POST -uri $authority -Body $ReqTokenBody
$Authentication = $result.access_token
$headers = @{"Authorization" = "Bearer "+ $Authentication}
#>


function getEmailAttachment($token){
    
    <#Param(
    [parameter(Mandatory=$true)]
    [DateTime]
    $date
    )#>

    $headers = @{"Authorization" = "Bearer $($token.access_token)" }
    # $lastdays = 30
    $date = (Get-Date -Date (get-date).AddDays(-$lastdays) -Format yyyy-MM-dd)
    $messages =$null
    $messages = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$MailAddress/messages?`$filter=IsRead eq false and hasattachments eq true and ReceivedDateTime ge $date" -Method GET -headers $headers 
    $messages.value |  Select-Object @{n="from";e={$_.from.emailaddress.address}},id,hasattachments,isread,flag,ccrepipeints,isdraft,@{n="sender";e={$_.sender.emailaddress.address}},receivedDateTime
    

    foreach ($message in $messages.value){ 
    
    # get attachments and save to file system
    $query = $url + "/" + $message.Id + "/attachments"
    $attachments = $null
    $attachments = Invoke-RestMethod -Uri $query -Method GET -headers $headers 

    do {
        # in case of multiple attachments in email
        foreach ($attachment in $attachments.value){
            
            $Name = $attachment.Name
            $Data = $attachment.contentBytes
            $attachmentData = [System.Convert]::FromBase64String($attachment.ContentBytes)
            
            #check if File exist
            $existiert = Test-Path -Path $OutputPath\$Name
            if ($existiert -eq $false)
                {
                $null = Set-Content -Path "$OutputPath\$Name" -Value $attachmentData -Encoding Byte
                    Write-Host "File '$Name' wurde angelegt" -Fore Green
                }
                else
                {
                    Write-Host "File '$Name' existiert schon." -Fore Red
                }
            

            Write-Host "ATTACHMENT: Name=$Name" -Fore Yellow

        }

        # set mail to IsRead = true
        $query2 = $url + "/" + $message.Id
        Invoke-RestMethod -Uri $query2 -Method Patch -headers $headers -ContentType "application/json" -Body '{"IsRead": true}'

        if ($attachments.'@Odata.NextLink') {
            $attachments = Invoke-RestMethod -Headers @{Authorization = "Bearer $($token.access_token)" } -Uri $attachments.'@Odata.NextLink' -Method "GET" -ContentType "application/json"
        }
    } while ($attachments.'@Odata.NextLink')
}
}


# Open Log
$prefix = $MyInvocation.MyCommand.Name
$stamp = (Get-Date).ToString().Replace("/", "-").replace(":", "-")
Start-Transcript -Path "C:\Users\gianl\Desktop\OneDrive_1_20.12.2020\log\log_final.txt" -Append
$start = Get-Date


#----- Entry Point -----#
main

# Close Log
$elapsed = (Get-Date) - $start
$elapsed
Stop-Transcript
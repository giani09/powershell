$OutputPath = "C:\tmp"
$lastdays = 10
$MailAddress = 'xxx' # Set the Mail Address
$tenantID = 'xxx' # Set the Tennant ID of the App Registration
$ClientID = 'xxx' # Set the Client ID of the App Registration
$CertificateThumbPrint = '0F7E69395AF02000A7734E16E4FD29E9ECEF84DB'
$url = "https://graph.microsoft.com/v1.0/users/$MailAddress/messages"

#Import-Module Azurerm
#Import-Module Msal.ps

$scope = "https://graph.microsoft.com/.default"

#----- Authenticates via OAuth -----#
function Authentication {
    try {
        $Authentication = Get-MsalToken -Scope $scope -TenantId $tenantID -ClientId $ClientID -ClientCertificate (Get-Item "cert:\LocalMachine\My\$CertificateThumbPrint" -ErrorAction Stop) -ErrorAction Stop         
        if($Authentication.AccessToken){
            writeOutput SUC ("Received Access Token.")
        }else{
            writeOutput ERR ("Did NOT receive Access Token.")
        }
    } catch {
        writeOutput ERR ("{0}" -f $Error[0])
    }
}

$headers = {Authorization= "Bearer "+ $Authentication.AccessToken.ToString()}
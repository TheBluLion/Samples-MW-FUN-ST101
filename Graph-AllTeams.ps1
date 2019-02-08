Function GetAccessToken
   {
    param (        
        [Parameter(Position=0, Mandatory=$false)] 
        [string] $Office365Username, 
        [Parameter(Position=1, Mandatory=$false)] 
        [string] $Office365Password
      )
    # Add ADAL (Microsoft.IdentityModel.Clients.ActiveDirectory.dll) assembly path from Azure Resource Manager SDK location
    #Add-Type -Path "C:\Program Files (x86)\Microsoft SDKs\Azure\PowerShell\ResourceManager\AzureResourceManager\AzureRM.Profile\Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
    # or simply import AzureRm module using below command
     Import-Module AzureRm
    #PowerShell Client Id. This is a well known Azure AD client id of PowerShell client. You don't need to create an Azure AD app.
    $clientId = "1950a258-227b-4e31-a9cf-717495945fc2"
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $resourceURI = "https://graph.microsoft.com"
    $authority = "https://login.microsoftonline.com/common"
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
     
    if (([string]::IsNullOrEmpty($Office365Username) -eq $false) -and ([string]::IsNullOrEmpty($Office365Password) -eq $false)) 
    { 
    $SecurePassword = ConvertTo-SecureString -AsPlainText $Office365Password -Force           
    #Build Azure AD credentials object  
    $AADCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential" -ArgumentList $Office365Username,$SecurePassword
    # Get token without login prompts.
    $authResult = $authContext.AcquireToken($resourceURI, $clientId,$AADCredential)
    } 
    else 
    {     
    # Get token by prompting login window.
    $authResult = $authContext.AcquireToken($resourceURI, $clientId, $redirectUri, [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Always)
    } 
    return $authResult.AccessToken
}

$token = GetAccessToken

$apiUrl = 'https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq ''Team'')'

$Teams = Invoke-RestMethod -Headers @{Authorization = "Bearer $token"} -Uri $apiUrl -Method Get
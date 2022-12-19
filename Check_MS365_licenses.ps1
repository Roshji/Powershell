<# 
.SYNOPSIS
    This script queries the Graph API to check for low availabilty of MS365 Licences.

.DESCRIPTION 
    This script queries the Graph API to check if there are still enough available licenses within the MS365 Tenant. If the availabilty is below x licenses it will send an e-mail with a warning. 

    This script must be used in combination with an Application Registration within Azure AD. The permissions needed for the App Registration are:

    Mail.Send - Application
    Organization.Read.All - Application
    User.Read - Delegated

    Variables that should be changed within the Script:
        $MailSender
        $MailReceiver
        $AppID
        $AppSecret
        $TenantID
        $Filteredobjects for SKU's you want to filter
 
.NOTES 
    Filename: Check_MS365_licenses.ps1
    Author: Remco van Diermen
    Version: 1.0

.COMPONENT 
    Built in Powershell v5.1

.LINK 
    https://www.remcovandiermen.nl
     
#>


Function AlertMail

{

#From which e-mailaddress is the mail sent from
$MailSender = "<SENDER MAIL ADDRESS>"

#From which e-mailaddress is the mail sent to
$MailReceiver = "<RECEIVER MAIL ADDRESS>"

$Attachment= "$env:TEMP\SKU.csv"
$FileName=(Get-Item -Path $Attachment).name
$base64string = [Convert]::ToBase64String([IO.File]::ReadAllBytes($Attachment))

#Connect to GRAPH API
$tokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $AppId
    Client_Secret = $AppSecret
}
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody
$headers = @{
    "Authorization" = "Bearer $($tokenResponse.access_token)"
    "Content-type"  = "application/json"
}

#Send Mail    
$URLsend = "https://graph.microsoft.com/v1.0/users/$MailSender/sendMail"
$BodyJsonsend = @"
                    {
                        "message": {
                          "subject": "You have Microsoft 365 License Warning(s)",
                          "body": {
                            "contentType": "HTML",
                            "content": "There are licensing warnings present in your Microsoft 365 Tenant. <br>
                            Please review the attachment <br>
                            
                            "
                          },
                          "toRecipients":[ 
                            {
                              "emailAddress": {
                                "address": "$MailReceiver"
                              }
                            }
                          ]
                        ,
                        "attachments": [
                            {
                              "@odata.type": "#microsoft.graph.fileAttachment",
                              "name": "$FileName",
                              "contentType": "text/plain",
                              "contentBytes": "$base64string"
                            }
                          ]
                        },
                        "saveToSentItems": "false"
                      }
"@

Invoke-RestMethod -Method POST -Uri $URLsend -Headers $headers -Body $BodyJsonsend
write-Output "Warnings Found, E-mail was sent"
}


# Define AppId, secret and scope, your tenant name and endpoint URL
$AppId = "<APP REGISTRATION ID>"
$TenantId = "<AZURE AD TENANT ID>"
$AppSecret = '<APP REGISTRATION SECRET>'
$Scope = "https://graph.microsoft.com/.default"

$Url = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"


# Add System.Web for urlencode
Add-Type -AssemblyName System.Web

# Create body
$Body = @{
    client_id = $AppId
	client_secret = $AppSecret
	scope = $Scope
	grant_type = 'client_credentials'
}

# Splat the parameters for Invoke-Restmethod for cleaner code
$PostSplat = @{
    ContentType = 'application/x-www-form-urlencoded'
    Method = 'POST'
    # Create string by joining bodylist with '&'
    Body = $Body
    Uri = $Url
}

# Request the token!
$Request = Invoke-RestMethod @PostSplat

# Create header
$Header = @{
    Authorization = "$($Request.token_type) $($Request.access_token)"
}

$Uri = "https://graph.microsoft.com/v1.0/subscribedSkus"

# Fetch all security alerts
$SKURequest = Invoke-RestMethod -Uri $Uri -Headers $Header -Method Get -ContentType "application/json"

$SKUS = $SKURequest.Value
$Report = [System.Collections.Generic.List[Object]]::new() 

Foreach ($SKU in $SKUS) {  

$FilteredObjects = @("VISIOCLIENT","PROJECTPREMIUM", "MCOMEETADV")
$CompareFilter = Compare-Object -ReferenceObject $SKU.skuPartNumber -DifferenceObject $FilteredObjects -IncludeEqual | where-object{$_.sideindicator -eq "=="}

    If (!$CompareFilter)

    {


    If (($SKU.capabilityStatus -ne "Enabled") -or ($SKU.consumedUnits -eq 0 )) 
        {
            write-output "Skipping" }
    Else 
        { 
         
            $ReportLine  = [PSCustomObject] @{          
        
            skuPartNumber = $SKU.skuPartNumber
            prepaidUnits = $SKU.prepaidUnits.enabled
            ComsumedUnits = $SKU.consumedUnits
            freeunits = ($SKU.prepaidUnits.enabled - $SKU.consumedUnits)
          }

    <# Add this part if you want to query on percentages

    $percentage = ($Reportline.freeunits / $reportline.prepaidUnits).tostring("P")

    If ([int]$percentage.trim("%") -le 10)
        {        
            $Report.Add($ReportLine) 
        }
    #>

    # Add this part if you want to query on values
   

    If ($Reportline.freeunits -le 10)
        {        
            $Report.Add($ReportLine) 
        }

    
        }

    }

}


#Create Report from array
$Report | Export-CSV "$env:TEMP\SKU.csv" -notypeinformation 

# Checking if there is a attachment to send via mail
$Attachment= "$env:TEMP\SKU.csv"
$QueryFile = Get-item -Path $Attachment -erroraction Silentlycontinue 

If ($QueryFile.length -gt 0)

    {
        AlertMail
    }
Else
    {
        exit
    }
function Get-AuthHeader {
    ##Tenant and App Specific Values
    ##Add the ones that you captured during the Azure portal piece here!
    $appID = "<<APP ID GOES HERE>>"
    $appSecret="<<APP SECRET KEY GOES HERE>>"
    ##Needs to be encoded so that special characters get passed through the URL correctly
    Add-Type -AssemblyName System.Web
    $appSecretEncoded = [System.Web.HttpUtility]::UrlEncode($appSecret)

    $tokenAuthURI = "<<OAUTH2.0 TOKEN ENDPOINT URL GOES HERE>>"

    ##We create a small text body with the values
    $requestBody = "grant_type=client_credentials" + 
        "&client_id=$appID" +
        "&client_secret=$appSecretEncoded" +
        "&resource=https://graph.microsoft.com/"

    ##Then we use the Token Endpoint URI and pass it the values in the body of the request
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenAuthURI -body $requestBody -ContentType "application/x-www-form-urlencoded"

    ##This response provides our Bearer Token
    $accessToken = $tokenResponse.access_token

    $AuthHeader = @{"Authorization"="Bearer $accessToken"}

    return $AuthHeader
}

function Get-GraphAPIReports {
    param (
        [Parameter(Mandatory = $true)]
        $AuthHeader,

        [Parameter(Mandatory = $true)]
        $LocalPath
    )

    $Period = "D180"
    $Date = (Get-Date -UFormat "%Y-%m-%d").ToString()
    $ReportArray = @("getEmailActivityUserDetail", "getEmailAppUsageUserDetail", "getMailboxUsageDetail","getOffice365GroupsActivityDetail", "getOneDriveActivityUserDetail", "getOneDriveUsageAccountDetail", "getSharePointActivityUserDetail", "getSharePointSiteUsageDetail", "getTeamsDeviceUsageUserDetail", "getTeamsUserActivityUserDetail","getYammerActivityUserDetail", "getYammerDeviceUsageUserDetail", "getYammerGroupsActivityDetail")

    #Loop through ReportsArray and put into LocalPath
    foreach ($Report in $ReportArray){
        #Build Parameter String
        $ParameterSet = $null

        #If period is specified then add that to the parameters
        if ($Period -and $Report -notlike "*Office365Activation*") {
            $Str = "period='{0}'," -f $Period
            $ParameterSet += $Str
        }
            
        #If the date is specified then add that to the parameters
        if ($Date -and !($Report -eq "MailboxUsage" -or $Report -notlike "*Office365Activation*" -or $Report -notlike "*getSkypeForBusinessOrganizerActivity*")) {
            $Str = "date='{0}'" -f $Date
            $ParameterSet += $Str
        }
        #Trim a trailing comma off the ParameterSet if needed
        if ($ParameterSet) {
            $ParameterSet = $ParameterSet.TrimEnd(",")
        }

        #Build the request URL and invoke
        $Body = @{
            output_mode = "csv"
        }

        try {
            $Uri = "https://graph.microsoft.com/beta/reports/{0}({1})" -f $Report, $ParameterSet
            $Result = Invoke-RestMethod -Uri $Uri -Headers $AuthHeader -Method Get -Body $Body 
            $ResultArray = ConvertFrom-Csv -InputObject $Result
            $ResultArray | Export-Csv -Path "$LocalPath\$Report.csv"
        }
        catch {
            write-host "Unable to get report $Report`nparameterset: $ParameterSet`nperiod: $Period`ndate: $Date`nuri: $Uri`nheaders: $AuthHeader`nbody: $Body`n"
        }
    }
}

function Get-UsersCustomAttributes {
    param
    (
        [Parameter(Mandatory = $true)]
        $LocalPath,

        [System.Management.Automation.PSCredential]
        $MyCredential
    )

    #Connect to Exchange Online
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $MyCredential -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
    Import-PSSession $Session -WarningAction SilentlyContinue -AllowClobber    

    $User_Info_Array = New-Object System.Collections.ArrayList
    $All_Users = Get-Mailbox -ResultSize Unlimited
    $j = 0

    foreach ($User in $All_Users) {
        $j++
        $UserArray = [pscustomobject]@{
            DisplayName = $User.DisplayName
            WindowsLiveID = $User.WindowsLiveID
            CustomAttribute8 = $User.CustomAttribute8
            CustomAttribute9 = $User.CustomAttribute9
        }

        $User_Info_Array.Add($UserArray) | Out-Null

        Write-Progress "Processing unified groups.." -Status "Processed $j of $($All_Users.count) groups" -PercentComplete (($j/$All_Users.count)*100) -CurrentOperation "$([System.Math]::Round(($j/$All_Users.count)*100,3))% complete"  
    }

    $User_Info_Array | Export-Csv "$LocalPath\UserInfoWithCustAtts.csv"

    $j = 0
    $User_Info_Array = $null
    $All_Users = $null
    $User = $null
}

function Send-SharePoint {
    param
    (
        [Parameter(Mandatory = $true)]
        $LocalPath,

        [System.Management.Automation.PSCredential]
        $MyCredential
    )

    $SPSiteURL = "https://tenant.sharepoint.com/sites/siteName"
    $SPDocLibName = "Documents" #Main doc library
    $SPFolderName = "O365_Reports" #Subfolder within main doc library

    #Change security protocol
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 

    #Add references to SharePoint client assemblies and authenticate to Office 365 site – required for CSOM
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

    #Bind to site collection
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SPSiteURL)
    $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($MyCredential.username, $MyCredential.password)
    $Context.Credentials = $Creds

    #Retrieve list
    $List = $Context.Web.Lists.GetByTitle($SPDocLibName)
    $Context.Load($List)
    $Context.ExecuteQuery()

    #Retrieve folder
    $FolderToBindTo = $List.RootFolder.Folders
    $Context.Load($FolderToBindTo)
    $Context.ExecuteQuery()
    $FolderToUpload = $FolderToBindTo | Where-Object {$_.Name -eq $SPFolderName}

    #Upload file
    Foreach ($File in (Get-ChildItem $LocalPath -File)) {
        $FileStream = New-Object IO.FileStream($File.FullName, [System.IO.FileMode]::Open)
        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
        $FileCreationInfo.Overwrite = $true
        $FileCreationInfo.ContentStream = $FileStream
        $FileCreationInfo.URL = $File
        $Upload = $FolderToUpload.Files.Add($FileCreationInfo)
        $Context.Load($Upload)
        $Context.ExecuteQuery()
    }
}

# Global Variables
$LocalPath = "C:\local\path"

#Get User Credentials
$User = "user@tenant.com"
$PasswordFile = "C:\Path\To\Encrypted\Password\File.txt"
$MyCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString)

# Get the authorization token
$AuthHeader = Get-AuthHeader -MyCredential $MyCredential

# Get the Graph API reports and save them to $Local Path
Get-GraphAPIReports -AuthHeader $AuthHeader -LocalPath $LocalPath

# Get List of Users from Exchange with Custom Attributes 8 (subcompany) & 9 (operating company)
Get-UsersCustomAttributes -LocalPath $LocalPath -MyCredential $MyCredential

# Send folder of reports ($LocalPath) to SharePoint
Send-SharePoint -LocalPath $LocalPath -MyCredential $MyCredential

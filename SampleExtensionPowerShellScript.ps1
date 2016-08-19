[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the URL of the target site, e.g. 'https://intranet.mydomain.com/sites/targetSite'")]
    [String]
    $TargetSiteUrl,

    [Parameter(Mandatory = $false, HelpMessage="Enter the filepath for the template, e.q. Folder\File.xml or Folder\File.pnp")]
    [String]
    $FilePath,

    [Parameter(Mandatory = $true, HelpMessage="Enter the comment to add!")]
    [String]
    $TemplateComment,

    [Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    [PSCredential]
    $Credentials
)

if(!$FilePath)
{
    $FilePath = "Extractions\site.xml"
}

if($Credentials -eq $null)
{
	$Credentials = Get-Credential -Message "Enter Admin Credentials"
}

Write-Host -ForegroundColor Yellow "Target Site URL: $targetSiteUrl"

try
{
    Connect-SPOnline $TargetSiteUrl -Credentials $Credentials -ErrorAction Stop

    [System.Reflection.Assembly]::LoadFrom("MyExtension\bin\Debug\MyExtension.dll") | Out-Null
    $commentIt = New-Object MyExtension.CommentIt
    $commentIt.Initialize($TemplateComment)

    Get-SPOProvisioningTemplate -Out $FilePath -Handlers Lists,Fields,ContentTypes,CustomActions -TemplateProviderExtensions $commentIt

    Disconnect-SPOnline

}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}
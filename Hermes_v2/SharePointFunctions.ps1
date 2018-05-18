$loadInfo1 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
$loadInfo2 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

Class DeliveryItem
{
    [int]$ID
    [String]$platform
    [String]$N16Version
    [Nullable[DateTime]]$CodeFreeze
    [String]$DevBranch
    [String]$ReleaseType
    [String]$TagBranch
    [bool]$HasBeenDelivered
}

function Get-AllDeliveries
{
  [CmdletBinding()]
  [OutputType([DeliveryItem[]])]
  param
  (
    [string]$siteUrl = 'https://nordealivogpension.edlund.dk/teamP/',
    [string]$listTitle = 'Deliveries'
  )
  
  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
  $credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
  $ctx.Credentials = $credentials
  $ctx.Load($ctx.Web)
  
  $lookupList = $ctx.Web.Lists.GetByTitle($listTitle)
  $ctx.Load($lookupList)

  $query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()

  $listItems = $lookupList.getItems($query)
  $ctx.Load($listItems)
  $ctx.ExecuteQuery()
  
  $listItems | % {
  New-Object DeliveryItem -Property @{
      ID = [int]$_['ID'];
      Platform = $_['Platform'];
      N16Version = $_['Title'];
      CodeFreeze = $_['Code_x002d_freeze'];
      DevBranch = $_['Dev_x002d_branch'];
      ReleaseType = $_['Release_x0020_type'];
      TagBranch = $_['Tag_x002d_branch'];
      HasBeenDelivered = $_['WasDelivered'];
    }
  }
}

function Get-UnbuiltDeliveries
{
  [CmdletBinding()]
  [OutputType([DeliveryItem[]])]
  param()

  Get-AllDeliveries | 
    ? { $_.HasBeenDelivered -eq $null -or $_.HasBeenDelivered -eq $False } 
}

function Get-NextDeliveryItem
{
  [CmdletBinding()]
  [OutputType([DeliveryItem])]
  param()

  Get-UnbuiltDeliveries | 
    ? { $_.CodeFreeze -ne $Null } |
    sort -Property 'CodeFreeze' |
    select -First 1
}


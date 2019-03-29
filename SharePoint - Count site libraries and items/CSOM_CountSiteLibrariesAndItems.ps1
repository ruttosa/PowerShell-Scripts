############################################################################################################################################ 
#Script that allows to create a new list in a SharePoint Online Site 
# Required Parameters: 
#  -> $sUserName: User Name to connect to the SharePoint Online Site Collection. 
#  -> $sPassword: Password for the user. 
#  -> $sCSOMPath: CSOM Assemblies Path. 
#  -> $sSiteUrl: SharePoint Online Site Url. 
#  -> $sListName: Name of the list we are going to create. 
#  -> $sListDescription: List description. 
############################################################################################################################################
### "¡·…ÈÕÌ”Û⁄˙¸‹Ò—#$%&/\()=¥¥+-{}[],;:.ø?°*~|∞"   Special Characters to quit in SharePoint Url's


##### SCRIPT PARAMETERS ######

#Param(
 #   [Parameter(Mandatory=$true)][string]$siteUrl, 
 #  [Parameter(Mandatory=$true)][string]$userName, 
 #   [Parameter(Mandatory=$true)][string]$pass
#)

##### SCRIPT PARAMETERS ######

######### FUNCTIONS ##########

## FunciÛn para escribir en el log de ERRORES de ejecuciÛn
Function LogWrite
{
   Param ([string]$logstring)
   $date = Get-Date
   $dateString = '[' + $date.ToString() + ']'
   $logStr = $dateString + ' ' + $logstring
   Add-content $Logfile -value $logStr
}
 

## FunciÛn para Contar Listas e Items de cada Biblioteca de un sitio
function countSiteLibrariesItems 
{
		
    $ListCollection = $spWeb.Lists
    $spCtx.Load($ListCollection)
    $spCtx.ExecuteQuery()
    $siteListsCount = $listCollection.Count
    
    $site_List_Items_Counter = 0
    if ($siteListsCount -gt 0){
        $index = 0
        do{
            $listObj = [Microsoft.SharePoint.Client.List] $ListCollection[$index]
            $list_items_counter = 0 
            $listTitle = $listObj.Title
            $listGuid = $listObj.Id
            
            $list = $spWeb.Lists.GetByTitle($listTitle)
            $spqQuery = New-Object Microsoft.SharePoint.Client.CamlQuery 
            $spqQuery.ViewXml =  
                                "<View>
                                    <RowLimit>2000</RowLimit>
                                    <ViewFields>                       
                                        <FieldRef Name='Id' />
                                    </ViewFields>
                                </View>" 
            do{
                $list_ItemsCollection = $list.getItems($spqQuery)
                $spCtx.Load($list_ItemsCollection)
                $spCtx.ExecuteQuery()
                $spqQuery.ListItemCollectionPosition = $list_ItemsCollection.ListItemCollectionPosition
                $list_items_counter += $list_ItemsCollection.Count
            }
            while($spqQuery.ListItemCollectionPosition -ne $null)

            Write-Host "# List:" $listTitle "-- Items:" $list_items_counter "-- GUID:" $listGuid -ForegroundColor Yellow
            $site_List_Items_Counter += $list_items_counter

            $index += 1
        }
        while($index -ne $siteListsCount)
    }
    Write-Host "Site Total Lists:" $siteListsCount "|| Total Items:" $site_List_Items_Counter -ForegroundColor Yellow
}

######### FUNCTIONS ##########


######## SCRIPT START ########

# Register Script Start Time
$scriptStartDate = Get-Date
Write-Host "Start: " + $scriptStartDate -ForegroundColor Yellow

# Global Script Configurations
$host.Runspace.ThreadOptions = "ReuseThread"
#$currentScriptPath = split-path -parent $MyInvocation.MyCommand.Definition  ### Get Script Current Path
#$currentScriptPathParent = split-path -parent $currentScriptPath
#$folderBase = $currentScriptPathParent + "\SharePoint DLL\" # Configure
#$logFileCreated = Get-Date -format 'yyyy_MM_dd_hhmm' 
#$Logfile = $currentScriptPath + '\' + $logFileCreated + "_log"

# Adding the Client Object Model Assemblies --> En caso de ser necesario    
$sCSOMRuntimePath=$folderBase +  "Microsoft.SharePoint.Client.Runtime.dll"         
$sCSOMPath=$folderBase +  "Microsoft.SharePoint.Client.dll"
Add-Type -Path $sCSOMPath          
Add-Type -Path $sCSOMRuntimePath

# Set Global Variables
$siteUrl = Read-Host "Site URL"

$sUserName = Read-Host "Usuario"
$Password = Read-Host "Password" -AsSecureString
$spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($sUserName, $Password)

# Verify Site Url Last Character
$url_last_char = $siteUrl.Substring($siteUrl.Length - 1, 1)
if($url_last_char -ne '/'){
    $siteUrl += '/'
}

# Set up global SPO Client Object Model Context variables
$spCtx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)   
$spCtx.Credentials = $spoCredentials
$spWeb = $spCtx.Web

# Register Script Function Execution Start Time
$scriptStartDate = Get-Date
Write-Host "Start: " + $scriptStartDate -ForegroundColor Yellow

# Specific Function Execution
countSiteLibrariesItems

Write-Host "End" -ForegroundColor Yellow 

######## SCRIPT END ########
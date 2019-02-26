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
### "¡·…ÈÕÌ”Û⁄˙¸‹Ò—#$%&/\()=¥¥+-{}[],;:.ø?°*~|∞"   Special Characters to quit

##### SCRIPT PARAMETERS ######

#Param(
#[Parameter(Mandatory=$true)][string]$userName, 
#[Parameter(Mandatory=$true)][string]$pass, 
#[Parameter(Mandatory=$true)][string]$siteUrl
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
 
## FunciÛn para Iniciar un flujo de trabajo sobre Items de cada Biblioteca de un sitio
function startWorkflowsOnListItems([string] $listName, [string] $workflowName)
{

    $list_items_total = 0
    $list_items_counter = 0

    # Instance Workflow Service Manager
    $serviceManager = New-Object Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager($spCtx, $spWeb)
    $spCtx.Load($serviceManager)

    ## Get Workflow Deployment Service
    $workflowDeploymentService = $serviceManager.GetWorkflowDeploymentService()
    $spCtx.Load($workflowDeploymentService)
    
    ## Get Published Workflows Collection
    $publishedWorkflowDefinitions = $workflowDeploymentService.EnumerateDefinitions($true)
    $spCtx.Load($publishedWorkflowDefinitions)
    $spCtx.ExecuteQuery()
    $wfDef_ID = ""

    ## Find Workflow Defintion ID
    foreach($wfDefintion in $publishedWorkflowDefinitions){
        $wfDef_Name = $wfDefintion.DisplayName
        if($wfDef_Name -eq $workflowName){
            $wfDef_ID = $wfDefintion.ID
        }
    }

    ## Get Subscriptions for Workflow
    $subscriptions = $serviceManager.GetWorkflowSubscriptionService().EnumerateSubscriptionsByDefinition($wfDef_ID)
    $spCtx.Load($subscriptions)
    $spCtx.ExecuteQuery()  

    ## Get List Items
    $ListOC = [Microsoft.SharePoint.Client.List] $spWeb.Lists.GetByTitle($listName)
    $spCtx.Load($ListOC)
    $spCtx.ExecuteQuery()

    $spqQuery = New-Object Microsoft.SharePoint.Client.CamlQuery 
    $spqQuery.ViewXml =  
                        "<View>
                            <RowLimit>2000</RowLimit>
                            <ViewFields>                       
                                <FieldRef Name='Id' />
                            </ViewFields>
                        </View>" 
    do{
        [Microsoft.SharePoint.Client.ListItemCollection] $ListOC_ItemsCollection = $ListOC.getItems($spqQuery)
        $spCtx.Load($ListOC_ItemsCollection)
        $spCtx.ExecuteQuery()
        $spqQuery.ListItemCollectionPosition = $ListOC_ItemsCollection.ListItemCollectionPosition
        $list_items_total += $ListOC_ItemsCollection.Count

        ## Process Each Item Found
        foreach($item in $ListOC_ItemsCollection){
            $item_ID = $item.Id
            Write-Host "Item ID:" $item_ID -ForegroundColor Yellow

            #### Block Code: Obtener el valor de un campo para activar el flujo de acuerdo a una condiciÛn que dependa de su valor ####
            #$itemOC = $ListOC.GetItemById($item_ID)
            #$spCtx.Load($itemOC)
            #$spCtx.ExecuteQuery()            
            #$item_status = $itemOC.FieldValues["Status"]
            #if($item_status -eq "No Iniciada"){
            $list_items_counter += 1
            
            ## Subscribe workflow on Item
            $subsEnum = $subscriptions.GetEnumerator()
            while($subsEnum.MoveNext()){
                $sub = $subsEnum.Current
                $initParam = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'
                $serviceManager.GetWorkflowInstanceService().StartWorkflowOnListItem($sub, $item_ID, $initParam)
                $spCtx.ExecuteQuery()
            }
            #}                                     
        }
    }
    while($spqQuery.ListItemCollectionPosition -ne $null)

    Write-Host "Total Items:" $list_items_counter -ForegroundColor Yellow
}


######### FUNCTIONS ##########

######## SCRIPT START ########

#Register Script Start Time
$scriptStartDate = Get-Date
Write-Host "Start: " + $scriptStartDate -ForegroundColor Yellow

#Global Script Configurations
$host.Runspace.ThreadOptions = "ReuseThread"
$currentScriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$parentCurrentPath = split-path -parent $currentScriptPath
$folderBase = $parentCurrentPath  + "\SharePoint DLL\"
$logFileCreated = Get-Date -format 'yyyy_MM_dd_hhmm' 
$Logfile = $currentScriptPath + '\' + $logFileCreated + "_log"

# Adding the Client Object Model Assemblies --> En caso de ser necesario    
$sCSOMRuntimePath=$folderBase +  "Microsoft.SharePoint.Client.Runtime.dll"         
$sCSOMPath=$folderBase +  "Microsoft.SharePoint.Client.dll"
$sCSOMPath2=$folderBase +  "Microsoft.SharePoint.Client.WorkflowServices.dll"
Add-Type -Path $sCSOMPath          
Add-Type -Path $sCSOMPath2 
Add-Type -Path $sCSOMRuntimePath   

# Set Global Variables
$siteUrl = Read-Host "Site URL"

$sUserName = Read-Host "Usuario"
$Password = Read-Host "Password" -AsSecureString
$spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($sUserName, $Password)

$sListName = Read-Host "Lista o Biblioteca"
$sWorkName = Read-Host "Flujo de trabajo"

# Verify Site Url Last Character
$url_last_char = $siteUrl.Substring($siteUrl.Length - 1, 1)
if($url_last_char -ne '/'){
    $siteUrl += '/'
}

# SPO Client Object Model Context
$spCtx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)   
$spCtx.Credentials = $spoCredentials
$spWeb = $spCtx.Web

startWorkflowsOnListItems -listName $sListName -workflowName $sWorkName

Write-Host "End" -ForegroundColor Yellow 

######## SCRIPT END ########
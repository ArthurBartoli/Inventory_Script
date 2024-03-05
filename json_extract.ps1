#! Report export
Login-PowerBIServiceAccount
Get-PowerBIWorkspace -Scope Organization -Include All -All | ConvertTo-Json -Depth 10 | Out-File "$PSScriptRoot\export.json"
#Invoke-PowerBIRestMethod -Url 'admin/groups?$top=5000&$skip=0&$expand=reports,dashboards,datasets,users,dataflows' -Method GET -Scope Organization  | Out-File "$PSScriptRoot\export.json"

#! Datasources export
# Obtenir tous les datasets
$datasets = Get-PowerBIDataset -Scope Organization
$results = @()

# Parcourir chaque dataset
foreach ($dataset in $datasets) {
    # Obtenir les sources de données pour le dataset actuel
    try {
        Write-Output $dataset.Name
        $dataSources = Get-PowerBIDatasource -DatasetId $dataset.Id -Scope Organization -ErrorAction Stop
        
        # Parcourir chaque source de données
        foreach ($dataSource in $dataSources) {
            $results += New-Object PSObject -Property @{
                DatasetId = $dataset.Id
                DatasetName = $dataset.Name
                DataSourceName = $dataSource.DataSourceName
                DataSourceId = $dataSource.DatasourceId
                DataSourceType = $dataSource.DatasourceType
                GatewayId = $dataSource.GatewayId 
                ConnectionDetails = $dataSource.ConnectionDetails
                ConnectionString = $dataSource.ConnectionString
            }
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($errorMessage -eq "Operation returned an invalid status code 'NotFound'") {
            Write-Output "### " + $errorMessage
            Write-Output "### Moving on to the next dataset."
        }
        else {
            Write-Output "An unexpected error has occured during the data source API query."
            Write-Output "Iteration over datasets will stop. It can either be because of an ID error or hitting the 50 query limit."
            Write-Output "Here is the error message :"
            Write-Output $errorMessage
        }
    }
}
$results | ConvertTo-Json | Out-File "$PSScriptRoot\DatasourceAndGateway.json"
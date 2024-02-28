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
# SIG # Begin signature block
# MIIFgQYJKoZIhvcNAQcCoIIFcjCCBW4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU0ZkVXh5X+R71SsCV6DfrifYR
# H/SgggMeMIIDGjCCAgKgAwIBAgIQQXoB6RffkLJNOhNZ2tZQajANBgkqhkiG9w0B
# AQsFADAYMRYwFAYDVQQDDA1BcnRodXJCYXJ0b2xpMB4XDTI0MDExNzEwMzcyNloX
# DTI1MDExNzEwNTcyNlowGDEWMBQGA1UEAwwNQXJ0aHVyQmFydG9saTCCASIwDQYJ
# KoZIhvcNAQEBBQADggEPADCCAQoCggEBAMu4egJZ9bnZi57w5rK5YR/2l2/7CY/f
# bLvSvgVVADSgiRYgCeJINbYjG9y1SnhP9Yj4y1SlMngExpmlcxTwVVj7WV2rrlhy
# S9OLz2Shs2r0pYL9z/YfvUN2xVYjwHUBOL88hou7maJ9wS3tIxoWLSLIkvsMt+5E
# X6KbJwFvz6aO8Q4UMvDfnTuKmcT9Y/8WR/vi5NKD0zlu9OfqbbCKYc0HSB5t1bQL
# gn1CH12J0kJa7VFQwOrxR1aei6rP0CTFd/OOg6a6D+TPgx/uuu2ulCs4gHh/GMyV
# J9nYPbaae7i6LV2AKvM/J37RZwqclT59V+kZ4nof2/CleCyF683g+rUCAwEAAaNg
# MF4wDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMBgGA1UdEQQR
# MA+CDUFydGh1ckJhcnRvbGkwHQYDVR0OBBYEFCenGXbSqgGaMjWWY31evODsYJXE
# MA0GCSqGSIb3DQEBCwUAA4IBAQAnHif87ZtSAbCv6r7aLk4HP436EDjz6lG9roIK
# OFX2EfZrTYoc2H2i14rcBXoTUaoOVy3IAP4r9eu6iN68zBk2J8ZHMbNiHRByfo82
# A7ZtnbTYwLw8DUqZ4DUj8YD/hFiKVVtBjaRDkHb134if1pCn5cGmuv/0Q50eoxr+
# 4mx6+3BGqQ0ZhJyDw4SyWFMKX7E6pMO5XPTo2UKfVZdvIz8P4lbr3ujyHdlIjfTt
# D2RspsVFTGrE62GwjawIa1b7N6YI1+wI5XZX1TBub4RAumXZ0wRXbRxQrE2KCL4c
# T3aeyqSH2Ex0UwVSu06uDaNeDwR6sZx4Vde0PU6HNBv393BwMYIBzTCCAckCAQEw
# LDAYMRYwFAYDVQQDDA1BcnRodXJCYXJ0b2xpAhBBegHpF9+Qsk06E1na1lBqMAkG
# BSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJ
# AzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMG
# CSqGSIb3DQEJBDEWBBQO7VI1mQDyAw3a764A9hhdAO7TDjANBgkqhkiG9w0BAQEF
# AASCAQBPmq6fE7SOcqlJA1SF/WvjRvAKHbPsxdrvgp5CoDiHNbSv7s/RUy3v/OmA
# Gok88TumXZl457d3XST33WaLrEnU8sBBbWS0ISI8negUbHeiCGPMFKFJzy7GjYtD
# pQol6+euR5ELFbN5UgZ7ArmiGlIRMfktZ4wKTWMuagJ8iPCW2MIOA17oV862KSJ5
# 8vAEtfM/pu0NpR2hM9L60GBs4wjEMTOjjC++JKJYqqLe0Er9oCFR3wnOi/UaWHWV
# psL5Rge3kIk3g96zX9GEUSRHdLzso4GCmfk/kKrejImf7c5trruPro4APVgZk9gX
# 9vnaWZzxl2mt7AshAkCNqJYDkaKq
# SIG # End signature block

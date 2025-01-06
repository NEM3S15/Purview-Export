<# 
Ce script a pour objectif de contourner la limitation du nombre de lignes extraites depuis l'Activity Explorer de Purview afin de faciliter l'analyse des événements DLP dans un fichier CSV formatté par colonnes
#>

# Utilisation de TLS 1.2 pour la session PowerShell en cours
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Utilisation de PowerShell derrière un proxy
(New-Object -TypeName System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

# Installation du module ExchangePowerShell et authentification
Install-Module -Name ExchangePowerShell
Connect-IPPSSession -UserPrincipalName example@domain.com

# Initialisation des variables
$scriptPath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptPath
$startTime = "MM/dd/yyyy"
$endTime = "MM/dd/yyyy"
$date = (Get-Date).ToString("ddMMyyyy_HHmm")
$watermark = $null
[string]$sourceDocument ="ActivityExplorerData-DLP-SPO-Teams-OD-EXO_$date.csv"
$path = "$dir\$sourceDocument"

# Premier export
$data = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -Filter1 @("Activity", "DLPRuleMatch") -Filter2 @("Workload", "OneDrive","SharePoint","MicrosoftTeams","Exchange") -Outputformat Json -PageSize 5000
$dataoutput = $data.resultdata
$watermark = $data.$WaterMark
$results = ConvertFrom-Json $dataoutput

# Le watermark reste vide si le résultat du premier export est inférieur à 5000 lignes
if ($watermark)
{
    do{
        $data = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -Filter1 @("Activity", "DLPRuleMatch") -Filter2 @("Workload", "OneDrive","SharePoint","MicrosoftTeams","Exchange") -Outputformat Json -PageSize 5000 -PageCookie $watermark
        $dataouput = $data.resultdata
        $resultoutput = ConvertFrom-Json $dataouput
        $results += $resultoutput
        $watermark = $data.$WaterMark
    } until (!$watermark)
}

foreach ($result in $results)
{
    # Construction du document avec les attributs à exporter
    $resultat = [PSCustomObject][Ordered] @{
        Activité = $result.Activity
        Chemin = $result.FilePath
        Workload = $result.Workload
        Utilisateur = ($result.User).ToLower()
        Domaine = (($result.User).ToLower() -Split ("@"))[1]
        Date = [datetime]::ParseExact($result.Happened, 'yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::CreateSpecificCulture('fr-FR'))
        Extension = $result.FileExtension
        Taille = $result.FileSize
        Low = [string]$result.SensitiveInfoTypeBucketsData.Low
        Medium = [string]$result.SensitiveInfoTypeBucketsData.Medium
        High = [string]$result.SensitiveInfoTypeBucketsData.High
        Total = $result.SensitiveInfoTypeData.SensitiveInformationDetectionsInfo.DetectedValues.Name.Count
        Match = (@($result.SensitiveInfoTypeData.SensitiveInformationDetectionsInfo.DetectedValues.Name) | Out-String).Trim()
        Stratégie = $result.PolicyMatchInfo.PolicyName
        Mode = $result.PolicyMatchInfo.PolicyMode
        Règle = $result.PolicyMatchInfo.RuleName
        Action = [array]$result.PolicyMatchInfo.RuleActions.Action -join ","
        Autre = [array]$result.PolicyMatchInfo.OtherConditions.Condition -join ","
    }
    $resultat | Sort-Object -Property Date | Export-Csv -Path $path -Encoding "UTF8" -NoTypeInformation -Append -delimiter ";"
}

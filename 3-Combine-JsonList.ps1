$DownloadDirectoryName = "Download"
$ExportDirectoryName = "Export"
$ProvidingText = "提供エリア内です"
$ProvidingPartiallyText = "提供エリア内です (要確認)"
$NotProvidingText = "お申込みいただけません"

Function Get-StateSummaryText ($StateSummary){
    Switch($StateSummary){
        "providing"{Return $ProvidingText}
        "providing_partially"{Return $ProvidingPartiallyText}
        "not_providing"{Return $NotProvidingText}
    }
}

Function Get-JsonData ($Path){
    $RawContent = (Get-Content $Path -Encoding UTF8 -Raw | ConvertFrom-Json).addresses
    ForEach ($Detail in $RawContent){
        $g2home = Get-StateSummaryText(($Detail.services | Where-Object category -eq "g2home").state_summary)
        $g2ms = Get-StateSummaryText(($Detail.services | Where-Object category -eq "g2ms").state_summary)
        $g10home = Get-StateSummaryText(($Detail.services | Where-Object category -eq "g10home").state_summary)
        $gs10home = Get-StateSummaryText(($Detail.services | Where-Object category -eq "gs10home").state_summary)
        #$Global:HashTable += $Detail | Select "zip","address_key","pref","city","town","region_code",@{Name="g2home";Expression={$g2home}},@{Name="g2ms";Expression={$g2ms}},@{Name="g10home";Expression={$g10home}},@{Name="gs10home";Expression={$gs10home}}
        $Global:HashTable += $Detail | Select @{Name="郵便番号";Expression={("〒" + ($_.zip).Insert(3,"-"))}},@{Name="市区町村";Expression={$_.city}},@{Name="町名";Expression={$_.town}},@{Name="NURO 光";Expression={$g2home}},@{Name="NURO 光 for マンション";Expression={$g2ms}},@{Name="NURO 光 10G";Expression={$g10home}},@{Name="NURO 光 10Gs・6Gs";Expression={$gs10home}}
    }
}


Get-ChildItem -Path $DownloadDirectoryName -Directory | ForEach-Object {
    $Progress = 0
    $PrefectureName = $_.Name
    $JsonFiles = Get-ChildItem -Path $_.FullName -Filter "*.json" -Recurse
    $Global:HashTable = @()
    $JsonFiles | ForEach-Object {
        #進捗
        Write-Progress -Activity $_.DirectoryName -Status $_ -PercentComplete ($Progress / ($JsonFiles.Count) * 100)
        $Progress++
        Get-JsonData $_.FullName
    }
    $Global:HashTable = $HashTable | Sort-Object "NURO 光", "郵便番号"
    $HashTable | Export-Csv -Path "$ExportDirectoryName\$PrefectureName.csv" -NoTypeInformation -Encoding UTF8
    Write-Host "Exported: $PrefectureName.csv"
    
    $LastProvidingText = $null
    $Global:SplitedHashTable = @()
    $SplitedHashTableCount = 1
    $SplitedHashTableIndex = 0
    $HashTableCount = 0
    $HashTable | ForEach-Object {
        $Global:SplitedHashTable += $_
        $SplitedHashTableIndex++
        $HashTableCount++

        If ($LastProvidingText -eq $null){$LastProvidingText = $_."NURO 光"}
        If ($LastProvidingText -ne $_."NURO 光"){
            $FileName = "$ExportDirectoryName\$PrefectureName.$LastProvidingText.$SplitedHashTableCount.csv"
            $SplitedHashTable | Export-Csv -Path $FileName -NoTypeInformation -Encoding UTF8
            Write-Host "Exported: $FileName"
            $SplitedHashTableCount = 1
            $SplitedHashTableIndex = 0
            $Global:SplitedHashTable = @()
            $LastProvidingText = $_."NURO 光"
        }
        If ($SplitedHashTableIndex -eq 1999){
            $FileName = "$ExportDirectoryName\$PrefectureName.$LastProvidingText.$SplitedHashTableCount.csv"
            $SplitedHashTable | Export-Csv -Path $FileName -NoTypeInformation -Encoding UTF8
            Write-Host "Exported: $FileName"
            $SplitedHashTableCount++
            $SplitedHashTableIndex = 0
            $Global:SplitedHashTable = @()
        }
        If ($HashTableCount -eq $HashTable.Count){
            $FileName = "$ExportDirectoryName\$PrefectureName.$LastProvidingText.$SplitedHashTableCount.csv"
            $SplitedHashTable | Export-Csv -Path $FileName -NoTypeInformation -Encoding UTF8
            Write-Host "Exported: $FileName"
        }
    }
}

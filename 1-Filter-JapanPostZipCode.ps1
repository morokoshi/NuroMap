$ZipCodeListFilePath = "ZipCodeList\KEN_ALL.CSV"
$ScriptTitle = "郵便番号の抽出"

If (!(Test-Path $ZipCodeListFilePath)){
    Invoke-WebRequest -Uri https://www.post.japanpost.jp/zipcode/dl/kogaki/zip/ken_all.zip -OutFile "ZipCodeList\ken_all.zip"
    Expand-Archive -Path "ZipCodeList\ken_all.zip" -DestinationPath "ZipCodeList\"
}

Function Check-Prompt($PromptTitle, $PromptMessage){
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","実行する"
    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No","実行しない"

    $PromptOptions = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
    $PromptResult = $Host.Ui.PromptForChoice($PromptTitle, $PromptMessage, $PromptOptions, 0)

    If ($PromptResult -eq 0){
        Return $True
    }

    Return $False
}

Function Choose-Mode{
    If (Check-Prompt "モードの選択" "市区町村ごとに選択しますか?")
    {
        CityMode
    }
    Else{
        Measure-Command {AllMode $PrefectureList $ZipCodeList}
    }
}

Function Initialize-PrefectureList{
    Clear-Host
    Write-Host "$ZipCodeListFilePath を読み込んでいます..."
    # 形式: https://www.post.japanpost.jp/zipcode/dl/readme.html
    $Global:ZipCodeList = Import-Csv -Path $ZipCodeListFilePath -Encoding Default -Header "全国地方公共団体コード","旧郵便番号","郵便番号","都道府県名 (半角カタカナ)","市区町村名 (半角カタカナ)","町域名 (半角カタカナ)","都道府県名","市区町村名","町域名","一町域が二以上の郵便番号で表される場合の表示","小字毎に番地が起番されている町域の表示","丁目を有する町域の場合の表示","一つの郵便番号で二以上の町域を表す場合の表示","更新の表示","変更理由"
    $Global:PrefectureList = $ZipCodeList | Select 都道府県名 | Get-Unique -AsString
}

Function Set-PrefectureName{
    Clear-Host
    $Global:PrefectureName = Read-Host -Prompt "都道府県名を入力"
    Clear-Host
    $Global:SuggestedPrefectureName = ($PrefectureList | Where-Object 都道府県名 -Like "*$PrefectureName*")[0].都道府県名

    If (Check-Prompt $ScriptTitle "都道府県名は $SuggestedPrefectureName でよろしいですか?"){
        $Global:PrefectureName = $SuggestedPrefectureName
    }
    Else{
        Set-PrefectureName
    }
}

Function Initialize-CityList{
    Clear-Host
    Write-Host "しばらくお待ちください..."
    $Global:CityList = $ZipCodeList | Where-Object 都道府県名 -eq $PrefectureName | Select 市区町村名 | Get-Unique -AsString
}

Function Set-CityName{
    Clear-Host
    $Global:CityName = Read-Host -Prompt "市町村名を入力"
    Clear-Host
    $Global:SuggestedCityName = ($CityList | Where-Object 市区町村名 -Like "*$CityName*")[0].市区町村名

    If (Check-Prompt $ScriptTitle "市区町村名は $SuggestedCityName でよろしいですか?"){
        $Global:CityName = $SuggestedCityName
    }
    Else{
        Set-CityName
    }
}

Function CityMode{
    Set-PrefectureName
    Initialize-CityList
    Set-CityName

    New-Item -ItemType Directory -Path "ZipCodeList\$PrefectureName" | Out-Null
    $ZipCodeList | Where-Object 市区町村名 -eq $CityName | Select 郵便番号 | Format-Table -HideTableHeaders | Out-File "ZipCodeList\$PrefectureName\$CityName.txt"
}

Function AllMode{
    Clear-Host
    Write-Host "しばらくお待ちください..."

    $PrefectureCurrentStatus = 0
    ForEach ($PrefectureName in $PrefectureList){
        $PrefectureName = $PrefectureName.都道府県名
        $CityList = $ZipCodeList | Where-Object 都道府県名 -eq $PrefectureName | Select 市区町村名 | Get-Unique -AsString
        New-Item -ItemType Directory -Path "ZipCodeList\$PrefectureName" | Out-Null

        $CityCurrentStatus = 0
        $CityList | ForEach-Object{
            $CityName = $_
            $CityName = $CityName.市区町村名
            
            $CityCurrentStatus++
            If ($PrefectureCurrentStatus -eq 0){
                $PercentComplete = ($CityCurrentStatus / $CityList.Count) / $PrefectureList.Count * 100
            }
            Else{
                $PercentComplete = ($CityCurrentStatus / $CityList.Count) / $PrefectureList.Count + ($PrefectureCurrentStatus / $PrefectureList.Count) * 100
            }
            If ($PercentComplete -ge 100){
                $PercentComplete = 100
            }
            Write-Progress -Activity "$PrefectureName の市町村を保存中" -Status $CityName -PercentComplete $PercentComplete

            $ZipCodeList | Where-Object 市区町村名 -eq $CityName | Select 郵便番号 | Format-Table -HideTableHeaders | Out-File "ZipCodeList\$PrefectureName\$CityName.txt"
        }
        $PrefectureCurrentStatus++
    }
}


If ($ZipCodeList.Count -eq 0){
    Initialize-PrefectureList
}
Choose-Mode


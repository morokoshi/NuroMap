# https://nuro.jp/hikari/service/area/?zip=500-0000

$Uri = "https://nuro.jp/hikari/area-wrapper/wrapper/area2/judge"
Function Get-NuroArea ($ZipCode, $DestinationPath){
    #$ZipCode2 = $ZipCode1.Insert(3, "-")
    $Body = @{category="g2home%2Cg2ms%2Cg10home%2Cgs10home"; zip=$ZipCode}
    Invoke-RestMethod -Uri $Uri -Method POST -Body $Body -OutFile $DestinationPath
    #Start-Sleep -Seconds 1.5
}

#郵便番号のリストを保存したフォルダー
$DirectoryName = "ZipCodeList"
$DirectoryName = (Resolve-Path $DirectoryName).Path
Get-ChildItem -Path $DirectoryName -Filter "*.txt" -Recurse | ForEach-Object {
    $Progress = 0
    $LocationName = ($_.FullName).Replace(".txt","")
    $LocationName = $LocationName.Replace("$DirectoryName\","")
    New-Item -ItemType Directory -Name "Download\$LocationName" -Force | Out-Null
    $Content = Get-Content $_.FullName
    
    ForEach ($ZipCode in $Content){
        #簡易Nullチェック
        If ($ZipCode){
            #進捗
            $Progress++
            Write-Progress -Activity $LocationName -Status $ZipCode -PercentComplete ($Progress / ($Content.Length) * 100)

            $DestinationPath = "Download\$LocationName\$ZipCode.json"
            If (!(Test-Path $DestinationPath)) {
                Get-NuroArea $ZipCode $DestinationPath
            }
        }
    }
}

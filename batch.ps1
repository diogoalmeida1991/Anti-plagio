Get-ChildItem "C:\Users\User\Documents\Sistemas Operacionais hash" -Recurse | ForEach-Object {
    $hash = Get-FileHash $_.FullName -Algorithm SHA256
    [PSCustomObject]@{
        Nome = $_.Name
        Caminho = $_.FullName
        Hash = $hash.Hash
    }
} | Export-Csv -Path "hashes.csv" -NoTypeInformation
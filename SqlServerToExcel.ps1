# Kết nối và xuất dữ liệu
$connectionString = "Your connection string"
$query = "select * from HU_CONTRACT_TYPE"
$Data = Invoke-Sqlcmd -ConnectionString $connectionString -Query $query
$Data | Export-Excel -Path "C:\Users\CongPC\Music\outputfile.xlsx"

# Write-Host "Hello, World!"

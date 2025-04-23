# Kết nối và xuất dữ liệu
$connectionString = "Password=tvc@123;User ID=sa;Initial Catalog=VANESA_LIVE;Data Source=192.168.60.22\SQLSERVER2022;TrustServerCertificate=True"
$query = "select * from HU_CONTRACT_TYPE"
$Data = Invoke-Sqlcmd -ConnectionString $connectionString -Query $query
$Data | Export-Excel -Path "C:\Users\CongPC\Music\outputfile.xlsx"

# Write-Host "Hello, World!"
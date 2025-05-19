# In ra màn hình
Write-Host "Hello, World!"

# Cài thư viện để SQL Server xuất ra Excel, hoặc để Excel import vào Sql Server
Install-Module -Name ImportExcel

# Lấy dữ liệu dưới DB
<pre># Thiết lập thông tin kết nối
$connectionString = "your connection string"

# Kết nối tới SQL Server
$connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
$connection.Open()

# Câu lệnh SELECT
$query = "SELECT * FROM HU_CONTRACT_TYPE"

# Tạo SqlCommand
$command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)

# Thực hiện câu lệnh và lấy dữ liệu
$reader = $command.ExecuteReader()

# Đọc dữ liệu và hiển thị
while ($reader.Read()) {
    # Giả sử bảng có cột CODE và NAME
    $code = $reader["CODE"]
    $name = $reader["NAME"]
    Write-Host "Code: $code, Name: $name"
}

# Đóng đối tượng reader và kết nối
$reader.Close()
$connection.Close()</pre>

# Gọi api mà có xác thực token bearer
curl -H "Authorization: Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhcHBUeXBlIjoiV0VCQVBQIiwic2lkIjoiN2RlNDc5M2ItZDAzOC00NjM3LTlkOWUtNTRkOGFlMzNkMzMwIiwidHlwIjoiY29uZ25jMDEiLCJpYXQiOiIwIiwiSXNBZG1pbiI6IlRydWUiLCJleHAiOjE3NDc3MTAxMjAsImlzcyI6Ikhpc3RhZmYgY3VzdG9tZXIiLCJhdWQiOiJIaXN0YWZmIGN1c3RvbWVyIn0.Ns1cdcB3LfS5HjefBiY9ZKkMX9Qq-BhSy6qGd_ZNiJk" https://localhost:44359/api/SysOrtherList/GetOtherListByType?typeCode=GENDER
<br>
<pre>curl -H "Authorization: Bearer YOUR_BEARER_TOKEN_HERE" https://api.example.com/endpoint</pre>

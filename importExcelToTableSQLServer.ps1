# Thiết lập thông tin kết nối
$connectionString = "Your connection string"

# Đường dẫn tới file Excel
$excelFilePath = "C:\Users\CongPC\Music\outputfile.xlsx"

# Đọc dữ liệu từ file Excel
$data = Import-Excel -Path $excelFilePath

# Kết nối tới SQL Server
$connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
$connection.Open()

# Lặp qua từng hàng và chèn vào bảng
foreach ($row in $data) {
    $query = "INSERT INTO HU_CONTRACT_TYPE (CODE, NAME, PERIOD, TYPE_ID, IS_BHXH, IS_BHYT, IS_BHTN, IS_BHTNLD_BNN, IS_ACTIVE) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5, @Value6, @Value7, @Value8, @Value9)"
    $command = $connection.CreateCommand()
    $command.CommandText = $query

    # Thêm tham số
    $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Value1", $row.CODE)))
    $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Value2", $row.NAME)))
    $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Value3", $row.PERIOD)))

    $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Value4", $row.TYPE_ID)))
    $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Value5", $row.IS_BHXH)))
    $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Value6", $row.IS_BHYT)))

    $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Value7", $row.IS_BHTN)))
    $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Value8", $row.IS_BHTNLD_BNN)))
    $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Value9", $row.IS_ACTIVE)))


    # In Log Để Kiểm Tra Giá Trị
    Write-Host "Inserting: $($row.CODE), $($row.NAME), $($row.PERIOD), $($row.TYPE_ID), $($row.IS_BHXH), $($row.IS_BHYT), $($row.IS_BHTN), $($row.IS_BHTNLD_BNN), $($row.IS_ACTIVE)"

    # Thực thi lệnh
    $command.ExecuteNonQuery()
}

# Đóng kết nối
$connection.Close()

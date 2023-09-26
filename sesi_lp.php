<!DOCTYPE html>
<html>
<head>
    <title>Tabel HTML dengan Colspan</title>
</head>
<body>
	<button id="exportButton">Export to Excel</button>
    <table class="table table-bordered">
        <thead>
            <tr>
                <th rowspan="3">No</th>
                <th rowspan="3">Kasus</th>
                <th colspan="16"><center>Berdasarkan Waktu</center></th>
                <th rowspan="3"><center>Jml</center></th>
            </tr>
            <tr>
                <?php
                // Koneksi ke database
                $servername = "localhost"; // Ganti dengan nama server Anda
                $username = "zidane"; // Ganti dengan username database Anda
                $password = "zidane"; // Ganti dengan password database Anda
                $database = "pasuruan_db"; // Ganti dengan nama database Anda

                $conn = new mysqli($servername, $username, $password, $database);

                // Periksa koneksi
                if ($conn->connect_error) {
                    die("Koneksi gagal: " . $conn->connect_error);
                }

                // Query SQL untuk mengambil semua data dari tabel tb_sasi
                $sql = "SELECT
                            SUBSTRING_INDEX(SUBSTRING_INDEX(nama, '(', -1), ' - ', 1) AS waktu_mulai,
                            SUBSTRING_INDEX(SUBSTRING_INDEX(nama, ' - ', -1), ')', 1) AS waktu_selesai
                        FROM
                            tb_sasi;";

                $result = $conn->query($sql);

                for ($i = 0; $i < $result->num_rows; $i++) {
                    $row = $result->fetch_assoc();
                    echo "<th colspan='2'>" . $row['waktu_mulai'] . "</th>";
                    // SIMPAN DATA WAKTU MULAI DAN WAKTU SELESAI KE DALAM ARRAY
                    $waktu_mulai[] = $row['waktu_mulai'];  
                }
                ?>
            </tr>
            <tr>
                <?php
                // Kembali ke awal hasil
                $result->data_seek(0);

                for ($i = 0; $i < $result->num_rows; $i++) {
                    $row = $result->fetch_assoc();
                    echo "<th colspan='2'>" . $row['waktu_selesai'] . "</th>";
                    // SIMPAN DATA WAKTU MULAI DAN WAKTU SELESAI KE DALAM ARRAY
                    $waktu_selesai[] = $row['waktu_selesai'];
                }
                ?>
            </tr>
            <!-- Kolom "Berdasarkan Waktu" -->
            <tr>
                <!-- Kolom kosong sudah dihapus -->
            </tr>
        </thead>
        <tbody>
            <!-- Kolom "No" dan "Kasus" -->
            <?php
            // Query SQL untuk mengambil data dari hasil query yang telah Anda berikan
            $query = "SELECT lp.id_sasi, sasi.nama AS nama_sasi, lp.id_tindak_pidana, tindak_pidana.nama AS nama_tindak_pidana, COUNT(*) AS total_count 
                      FROM tb_lp lp 
                      JOIN tb_sasi sasi ON lp.id_sasi = sasi.id 
                      JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                      GROUP BY lp.id_sasi, lp.id_tindak_pidana";

            $result = $conn->query($query);

            $no = 1; // Untuk nomor urut
            $totalKanan = 0; // Initialize the totalKanan variable
            while ($row = $result->fetch_assoc()) {
                echo "<tr>";
                echo "<td>" . $no . "</td>";
                // Isi nama kasus dengan nama_tindak_pidana
                echo "<td>" . $row['nama_tindak_pidana'] . "</td>";
                // Tambahkan loop untuk mengisi data waktu dan jumlah
                for ($i = 0; $i < count($waktu_mulai); $i++) {
                    // Query SQL untuk mengambil data dari hasil query yang telah Anda berikan
                    $query = "SELECT lp.id_sasi, sasi.nama AS nama_sasi, lp.id_tindak_pidana, tindak_pidana.nama AS nama_tindak_pidana, COUNT(*) AS total_count 
                              FROM tb_lp lp 
                              JOIN tb_sasi sasi ON lp.id_sasi = sasi.id 
                              JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                              WHERE lp.id_sasi = '" . $row['id_sasi'] . "' AND lp.id_tindak_pidana = '" . $row['id_tindak_pidana'] . "' AND sasi.nama LIKE '%" . $waktu_mulai[$i] . "%' AND sasi.nama LIKE '%" . $waktu_selesai[$i] . "%'";

                    $result2 = $conn->query($query);
                    $row2 = $result2->fetch_assoc();
                    echo "<td colspan=2>" . $row2['total_count'] . "</td>";
                }
                echo "<td>" . $row['total_count'] . "</td>";
                echo "</tr>";
                $no++; // Tambah nomor urut
                $totalKanan += $row['total_count']; // Add the current row's total_count to totalKanan
            }
            echo "<tr>";
            echo "<td colspan=2><center>Jumlah</center></td>";
            for ($i = 0; $i < count($waktu_mulai); $i++) {
                // Query SQL untuk mengambil data dari hasil query yang telah Anda berikan
                $query = "SELECT COUNT(*) AS total_count 
                          FROM tb_lp lp 
                          JOIN tb_sasi sasi ON lp.id_sasi = sasi.id 
                          JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                          WHERE sasi.nama LIKE '%" . $waktu_mulai[$i] . "%' AND sasi.nama LIKE '%" . $waktu_selesai[$i] . "%'";

                $result2 = $conn->query($query);
                $row2 = $result2->fetch_assoc();
                echo "<td colspan=2>" . $row2['total_count'] . "</td>";
                $totalKanan += $row2['total_count']; // Add the current row's total_count to totalKanan
            }
            echo "<td>" . $totalKanan . "</td>"; // Display the totalKanan in the "Total Kanan" cell
            echo "</tr>";
            ?>
        </tbody>
    </table>
</body>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.2.0/jquery.min.js"></script>
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" />
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js"></script>
    <script>
        function exportToExcel() {
            const table = document.querySelector('table');
            const wb = XLSX.utils.table_to_book(table, { sheet: 'Sheet JS' });

            // Adjust the sheet content to account for rowspan and colspan
            const sheet = wb.Sheets['Sheet JS'];
            const mergeCells = [
                // Define your merged cells here as an array of objects.
                // For example, if you have a rowspan of 3 in the first cell, you would define it as:
                // { s: { r: 0, c: 0 }, e: { r: 2, c: 0 } }
            ];

            // Apply merged cells to the sheet
            sheet['!merges'] = mergeCells;
			var namefiletime = new Date().getTime();
			// Export the sheet to Excel file
            XLSX.writeFile(wb, 'LP Sasi ' + namefiletime + '.xlsx');
        }

        document.getElementById('exportButton').addEventListener('click', exportToExcel);
    </script>
</html>
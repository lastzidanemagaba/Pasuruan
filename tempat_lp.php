<!DOCTYPE html>
<html>
<head>
    <title>Tabel HTML dengan Colspan</title>
    <style>
        table {
            width: 80%; /* Adjust the table width as needed */
            font-size: 14px; /* Adjust the font size as needed */
            margin: 0 auto; /* Center the table horizontally */
        }
        table, th, td {
            border: 1px solid black;
            border-collapse: collapse;
        }
    </style>
</head>
<body>
    <button id="exportButton">Export to Excel</button>
    <table class="table table-bordered">
        <thead>
            <tr>
				<th rowspan="3">No</th>
                <th rowspan="3">Kasus</th>
                <th colspan="18"><center>Berdasarkan Tempat</center></th>
				<th rowspan="3">Jml</th>
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
                $sql = "SELECT *
				FROM
					tb_tempat;";

                $result = $conn->query($sql);

                for ($i = 0; $i < $result->num_rows; $i++) {
                    $row = $result->fetch_assoc();
                    echo "<th colspan='2'>" . $row['nama'] . "</th>";
                    // SIMPAN DATA WAKTU MULAI DAN WAKTU SELESAI KE DALAM ARRAY
                    $tempat[] = $row['nama'];
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
            $getID = $_GET['id'];
            if ($getID == "all") {
                $query = "SELECT DISTINCT tindak_pidana.nama AS nama_tindak_pidana 
                          FROM tb_lp lp 
					      JOIN tb_tempat tempat ON lp.id_tempat = tempat.id 
                          JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id
                          JOIN tb_wilayah wilayah ON lp.id_wilayah = wilayah.id;";
            }
            else if ($getID == "") {
                echo "Masukkan ID Wilayah terlebih dahulu";
            }
            else {
            $query = "SELECT DISTINCT tindak_pidana.nama AS nama_tindak_pidana 
                      FROM tb_lp lp 
					  JOIN tb_tempat tempat ON lp.id_tempat = tempat.id 
                      JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id
                      JOIN tb_wilayah wilayah ON lp.id_wilayah = wilayah.id 
					  WHERE wilayah.id = $getID";
            }

            $result = $conn->query($query);

            $no = 1; // Untuk nomor urut
            $totalKanan2 = 0;
			$totalKanan3 = 0; // Initialize the totalKanan variable
            while ($row = $result->fetch_assoc()) {
                echo "<tr>";
                echo "<td>" . $no . "</td>";
                // Isi nama kasus dengan nama_tindak_pidana
                echo "<td>" . $row['nama_tindak_pidana'] . "</td>";
                // Tambahkan loop untuk mengisi data waktu dan jumlah
                for ($i = 0; $i < count($tempat); $i++) {
                    if ($getID == "all") {
                        // Query SQL untuk mengambil data dari hasil query yang telah Anda berikan
                        $query = "SELECT COUNT(*) AS total_count 
                                  FROM tb_lp lp 
                                  JOIN tb_tempat tempat ON lp.id_tempat = tempat.id 
                                  JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                                  WHERE tindak_pidana.nama = '" . $row['nama_tindak_pidana'] . "' AND tempat.nama LIKE '%" . $tempat[$i] . "%';";
                    }
                    else if ($getID == "") {
                        echo "Masukkan ID Wilayah terlebih dahulu";
                    }
                    else {
                    $query = "SELECT COUNT(*) AS total_count 
                              FROM tb_lp lp 
                              JOIN tb_tempat tempat ON lp.id_tempat = tempat.id 
                              JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                              JOIN tb_wilayah wilayah ON lp.id_wilayah = wilayah.id 
                              WHERE tindak_pidana.nama = '" . $row['nama_tindak_pidana'] . "' AND tempat.nama LIKE '%" . $tempat[$i] . "%' AND wilayah.id = $getID;";
                    }
                    $result2 = $conn->query($query);
                    $row2 = $result2->fetch_assoc();
                    echo "<td colspan=2>" . $row2['total_count'] . "</td>";
                }
                // sum total
                if ($getID == "all") {
                    $query2 = "SELECT COUNT(*) AS total_count 
                              FROM tb_lp lp 
                              JOIN tb_tempat tempat ON lp.id_tempat = tempat.id 
                              JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                              WHERE tindak_pidana.nama = '" . $row['nama_tindak_pidana'] . "'";
                }
                else if ($getID == "") {
                    echo "Masukkan ID Wilayah terlebih dahulu";
                }
                else {
				$query2 = "SELECT COUNT(*) AS total_count 
						  FROM tb_lp lp 
						  JOIN tb_tempat tempat ON lp.id_tempat = tempat.id 
						  JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                          JOIN tb_wilayah wilayah ON lp.id_wilayah = wilayah.id 
						  WHERE tindak_pidana.nama = '" . $row['nama_tindak_pidana'] . "' AND wilayah.id = $getID";
                }
				$result3 = $conn->query($query2);
				$row3 = $result3->fetch_assoc();
				echo "<td>" . $row3['total_count'] . "</td>";


                echo "</tr>";
                $no++; // Tambah nomor urut
                $totalKanan2 += $row2['total_count']; // Add the current row's total_count to totalKanan
            }
            echo "<tr>";
            echo "<td colspan=2><center>Jumlah</center></td>";
            for ($i = 0; $i < count($tempat); $i++) {
                if ($getID == "all") {
                    // Query SQL untuk mengambil data dari hasil query yang telah Anda berikan
                    $query = "SELECT COUNT(*) AS total_count 
                              FROM tb_lp lp 
                              JOIN tb_tempat tempat ON lp.id_tempat = tempat.id 
                              JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                              WHERE tempat.nama LIKE '%" . $tempat[$i] . "%';";
                }
                else if ($getID == "") {
                    echo "Masukkan ID Wilayah terlebih dahulu";
                }
                else {
                $query = "SELECT COUNT(*) AS total_count 
                          FROM tb_lp lp 
                          JOIN tb_tempat tempat ON lp.id_tempat = tempat.id 
                          JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                          JOIN tb_wilayah wilayah ON lp.id_wilayah = wilayah.id 
                        WHERE tempat.nama LIKE '%" . $tempat[$i] . "%' AND wilayah.id = $getID;";
                }

                $result2 = $conn->query($query);
                $row2 = $result2->fetch_assoc();
                echo "<td colspan=2>" . $row2['total_count'] . "</td>";
                $totalKanan3 += $row2['total_count']; // Add the current row's total_count to totalKanan
            }
            echo "<td>" . $totalKanan3 . "</td>"; // Display the totalKanan in the "Total Kanan" cell
            echo "</tr>";
            ?>
        </tbody>
    </table>
</body>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.2.0/jquery.min.js"></script>
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

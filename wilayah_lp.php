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
                <th colspan="56"><center>Berdasarkan Wilayah</center></th>
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
						tb_wilayah;";

                $result = $conn->query($sql);

                for ($i = 0; $i < $result->num_rows; $i++) {
                    $row = $result->fetch_assoc();
					echo "<th colspan='2'>" . $row['nama'] . "</th>";
					$tempat[] = $row['nama'];
                }
                ?>
            </tr>
        </thead>
        <tbody>
        <!-- Kolom "No" dan "Kasus" -->
		<?php
            // Query SQL untuk mengambil data dari hasil query yang telah Anda berikan
            $query = "SELECT tindak_pidana.nama AS nama_tindak_pidana,tempat.nama as nama_tempat,COUNT(*) AS total_count 
			from tb_lp lp 
			JOIN tb_wilayah tempat ON lp.id_tempat = tempat.id 
			JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
			group by lp.id_tempat, lp.id_tindak_pidana;";

            $result = $conn->query($query);

			$no = 1; // Untuk nomor urut
            $totalKanan = 0; // Initialize the totalKanan variable
            while ($row = $result->fetch_assoc()) {
                echo "<tr>";
                echo "<td>" . $no . "</td>";
                // Isi nama kasus dengan nama_tindak_pidana
                echo "<td>" . $row['nama_tindak_pidana'] . "</td>";
                // Tambahkan loop untuk mengisi data waktu dan jumlah
                for ($i = 0; $i < count($tempat); $i++) {
					// Query SQL untuk mengambil data jumlah
					$query = "SELECT COUNT(*) AS total_count 
							  FROM tb_lp lp 
							  JOIN tb_wilayah tempat ON lp.id_tempat = tempat.id 
							  JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
							  WHERE tindak_pidana.nama = '" . $row['nama_tindak_pidana'] . "' AND tempat.nama = '" . $tempat[$i] . "';";

					$result2 = $conn->query($query);
					$row2 = $result2->fetch_assoc();
					echo "<td colspan=2>" . $row2['total_count'] . "</td>";
					$totalKanan += $row2['total_count'];
				}
				echo "<td>" . $totalKanan . "</td>";
				$totalKanan = 0;
				echo "</tr>";
				$no++;
			}
			echo "<tr>";
            echo "<td colspan=2><center>Jumlah</center></td>";
			for ($i = 0; $i < count($tempat); $i++) {
				// Query SQL untuk mengambil data dari hasil query yang telah Anda berikan
				$query = "SELECT COUNT(*) AS total_count 
						  FROM tb_lp lp 
						  JOIN tb_wilayah tempat ON lp.id_tempat = tempat.id 
						  JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
						  WHERE tempat.nama = '" . $tempat[$i] . "';";

				$result2 = $conn->query($query);
				$row2 = $result2->fetch_assoc();
				echo "<td colspan=2>" . $row2['total_count'] . "</td>";
				$totalKanan += $row2['total_count']; // Add the current row's total_count to totalKanan
			}
			echo "<td>" . $totalKanan . "</td>";
			echo "</tr>";

            ?>
        </tbody>
    </table>
</body>
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
                // Merge the "No" and "Kasus" header cells
				{ s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
				{ s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
				// Merge the "Berdasarkan Tempat" header cells
				{ s: { r: 0, c: 2 }, e: { r: 0, c: 19 } },
				// Merge the "Jumlah" header cells
				{ s: { r: 0, c: 20 }, e: { r: 1, c: 20 } },
				// Merge the "Jumlah" footer cells
				{ s: { r: 2, c: 20 }, e: { r: 3, c: 20 } },
			];


            // Apply merged cells to the sheet
            sheet['!merges'] = mergeCells;

			// Generate a unique filename
			const namefiletime = new Date().getTime();
			
			// Export the sheet to an Excel file with a unique filename
            XLSX.writeFile(wb, 'LP Wilayah ' + namefiletime + '.xlsx');
        }

        document.getElementById('exportButton').addEventListener('click', exportToExcel);
    </script>
</html>

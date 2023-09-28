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
				<th>No</th>
                <th>Kasus</th>
				<th>Jml</th>
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
            $getID = $_GET['id'];
            if ($getID == "all") {
                $query = "SELECT tindak_pidana.nama AS nama_tindak_pidana , 
                COUNT(*) as total_count 
                FROM tb_lp lp 
                JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                JOIN tb_wilayah wilayah ON lp.id_wilayah = wilayah.id 
                GROUP BY lp.id_tindak_pidana;";

            } else if ($getID == "") {
                echo "Masukkan ID Wilayah terlebih dahulu";
            } else {
                $query = "SELECT tindak_pidana.nama AS nama_tindak_pidana , 
                COUNT(*) as total_count 
                FROM tb_lp lp 
                JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                JOIN tb_wilayah wilayah ON lp.id_wilayah = wilayah.id 
                WHERE wilayah.id = $getID
                GROUP BY lp.id_tindak_pidana;";
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
                // Isi jumlah dengan total_count
                echo "<td>" . $row['total_count'] . "</td>";
                echo "</tr>";
                $no++;
                $totalKanan2 += $row['total_count'];
            }
            echo "<tr>";
            echo "<td colspan=2><center>Jumlah</center></td>";
            echo "<td>" . $totalKanan2 . "</td>";
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

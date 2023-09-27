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
                <th colspan="14"><center>Berdasarkan Hari</center></th>
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

                // Query SQL untuk mengambil semua data dari tabel tb_hari
                $sql = "SELECT *
                        FROM tb_hari;";

                $result = $conn->query($sql);

                $hariColumn = []; // Array to store day column names

                // Loop to display day column headers
                for ($i = 0; $i < $result->num_rows; $i++) {
                    $row = $result->fetch_assoc();
                    echo "<th colspan='2'>" . $row['nama'] . "</th>";
                    $hariColumn[] = $row['nama']; // Store day column names in the array
                }
                ?>
            </tr>
        </thead>
        <tbody>
            <?php
            // Query SQL to retrieve data
            $query = "SELECT tindak_pidana.nama AS nama_tindak_pidana, 
                DAYNAME(lp.tanggal) AS nama_hari, 
                COUNT(*) AS total_count 
                FROM `tb_lp` lp 
                JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                GROUP BY DAYNAME(lp.tanggal), lp.id_tindak_pidana;";

            $result = $conn->query($query);

            $no = 1; // Untuk nomor urut

            // Initialize total count variables
            $totalKanan = 0;
            $totalJumlah = 0;

            // Initialize an array to store data
            $data = [];

            // Loop to fetch and organize data
            while ($row = $result->fetch_assoc()) {
                $tindakPidana = $row['nama_tindak_pidana'];
                $hari = $row['nama_hari'];
                $total = $row['total_count'];

                // Store data in the array
                if (!isset($data[$tindakPidana])) {
                    $data[$tindakPidana] = [];
                }

                $data[$tindakPidana][$hari] = $total;
            }

            // Loop to display data in the table
            foreach ($data as $tindakPidana => $hariData) {
                echo "<tr>";
                echo "<td>" . $no . "</td>";
                echo "<td>" . $tindakPidana . "</td>";

                // Loop to display the counts for each day column
                foreach ($hariColumn as $hari) {
                    if (isset($hariData[$hari])) {
                        echo "<td colspan='2'>" . $hariData[$hari] . "</td>";
                        $totalKanan += $hariData[$hari];
                    } else {
                        echo "<td colspan='2'>0</td>";
                    }
                }

                echo "<td>" . $totalKanan . "</td>";
                echo "</tr>";
                $no++;
                $totalKanan = 0; // Reset the total count for the next row
            }

            // Display the "Jumlah" row for each day column
            echo "<tr>";
            echo "<td colspan='2'><center>Jumlah</center></td>";

            // Calculate and display the total count for each day column
            foreach ($hariColumn as $hari) {
                $totalHari = 0;
                foreach ($data as $tindakPidana => $hariData) {
                    if (isset($hariData[$hari])) {
                        $totalHari += $hariData[$hari];
                    }
                }
                echo "<td colspan='2'>" . $totalHari . "</td>";
                $totalJumlah += $totalHari;
            }

            echo "<td>" . $totalJumlah . "</td>";
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
            // Merge the "No" and "Kasus" header cells
            { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
            { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
            // Merge the "Berdasarkan Hari" header cells
            { s: { r: 0, c: 2 }, e: { r: 0, c: 15 } },
            // Merge the "Jml" header cells
            { s: { r: 0, c: 16 }, e: { r: 1, c: 16 } },
            // Merge the "Jumlah" footer cells
            { s: { r: 2, c: 16 }, e: { r: 3, c: 16 } },
        ];

        // Apply merged cells to the sheet
        sheet['!merges'] = mergeCells;

        // Generate a unique filename
        const namefiletime = new Date().getTime();

        // Export the sheet to an Excel file with a unique filename
        XLSX.writeFile(wb, 'LP Hari ' + namefiletime + '.xlsx');
    }

    document.getElementById('exportButton').addEventListener('click', exportToExcel);
</script>
</html>

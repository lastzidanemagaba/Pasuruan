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
                <th colspan="10"><center>Berdasarkan Usia Tersangka</center></th>
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
                        tb_usia_pelaku;";

                $result = $conn->query($sql);

                $usia_pelaku = array(); // Create an array to store usia_pelaku values

                for ($i = 0; $i < $result->num_rows; $i++) {
                    $row = $result->fetch_assoc();
                    echo "<th colspan='2'>" . $row['nama'] . "</th>";
                    $usia_pelaku[] = $row['nama'];
                }
                ?>
            </tr>
        </thead>
        <tbody>
        <!-- Kolom "No" dan "Kasus" -->
        <?php
        $getID = $_GET['id'];
        if ($getID == "all") {
            $queryTindakPidana = "SELECT DISTINCT tindak_pidana.nama AS nama_tindak_pidana 
                            FROM tb_tersangka korban
                            JOIN tb_lp lp ON korban.id = lp.id
                            JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id;";
        } else if ($getID == "") {
            echo "Masukkan ID Wilayah terlebih dahulu";
        } else {
        $queryTindakPidana = "SELECT DISTINCT tindak_pidana.nama AS nama_tindak_pidana 
                            FROM tb_tersangka korban
                            JOIN tb_lp lp ON korban.id = lp.id
                            JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id
                            JOIN tb_wilayah wilayah ON lp.id_wilayah = wilayah.id 
                            WHERE wilayah.id = $getID;";
        }
        $resultTindakPidana = $conn->query($queryTindakPidana);

        $no = 1; // Untuk nomor urut
        $totalKanan = 0; // Initialize the totalKanan variable

        while ($rowTindakPidana = $resultTindakPidana->fetch_assoc()) {
            echo "<tr>";
            echo "<td>" . $no . "</td>";
            // Isi nama kasus dengan nama_tindak_pidana
            echo "<td>" . $rowTindakPidana['nama_tindak_pidana'] . "</td>";
            // Initialize an array to store counts for each usia_pelaku
            $usia_pelakuCounts = array();

            // Query SQL to get counts for each usia_pelaku
            foreach ($usia_pelaku as $usia) {
                if ($getID == "all") {
                    $query = "SELECT COUNT(*) AS total_count 
                    FROM tb_tersangka korban 
                    JOIN tb_lp lp ON korban.id = lp.id 
                    JOIN tb_usia_pelaku usia_pelaku ON korban.id_usia_pelaku = usia_pelaku.id 
                    JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                    WHERE usia_pelaku.nama = '" . $usia . "' AND tindak_pidana.nama = '" . $rowTindakPidana['nama_tindak_pidana'] . "';";
                } else if ($getID == "") {
                    echo "Masukkan ID Wilayah terlebih dahulu";
                } else {
                $query = "SELECT COUNT(*) AS total_count 
                          FROM tb_tersangka korban 
                          JOIN tb_lp lp ON korban.id = lp.id 
                          JOIN tb_usia_pelaku usia_pelaku ON korban.id_usia_pelaku = usia_pelaku.id 
                          JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                          JOIN tb_wilayah wilayah ON lp.id_wilayah = wilayah.id 
                          WHERE usia_pelaku.nama = '" . $usia . "' AND tindak_pidana.nama = '" . $rowTindakPidana['nama_tindak_pidana'] . "' AND wilayah.id = $getID;";
                }
                $result2 = $conn->query($query);
                $row2 = $result2->fetch_assoc();

                // Store the count in the usia_pelakuCounts array
                $usia_pelakuCounts[] = $row2['total_count'];
            }

            // Output the counts for each usia_pelaku
            foreach ($usia_pelakuCounts as $count) {
                echo "<td colspan=2>" . $count . "</td>";
            }

            // Calculate and output the total for this tindak_pidana
            $totalTindakPidana = array_sum($usia_pelakuCounts);
            echo "<td>" . $totalTindakPidana . "</td>";

            // Reset the usia_pelakuCounts array
            $usia_pelakuCounts = array();

            echo "</tr>";
            $no++;
        }
        ?>
        <tr>
            <td colspan="2"><center>Jumlah</center></td>
            <?php
            // Loop to calculate and output the total for each usia_pelaku
            foreach ($usia_pelaku as $usia) {
                if ($getID == "all") {
                    $query = "SELECT COUNT(*) AS total_count 
                    FROM tb_tersangka korban 
                    JOIN tb_lp lp ON korban.id = lp.id 
                    JOIN tb_usia_pelaku usia_pelaku ON korban.id_usia_pelaku = usia_pelaku.id 
                    JOIN tb_tindak_pidana tindak_pidana ON lp.id_tindak_pidana = tindak_pidana.id 
                    WHERE usia_pelaku.nama = '" . $usia . "';";
                } else if ($getID == "") {
                    echo "Masukkan ID Wilayah terlebih dahulu";
                } else {
                $query = "SELECT COUNT(*) AS total_count 
                          FROM tb_tersangka korban 
                          JOIN tb_lp lp ON korban.id = lp.id 
                          JOIN tb_usia_pelaku usia_pelaku ON korban.id_usia_pelaku = usia_pelaku.id 
                          JOIN tb_wilayah wilayah ON lp.id_wilayah = wilayah.id 
                          WHERE usia_pelaku.nama = '" . $usia . "' AND wilayah.id = $getID;";
                }

                $result2 = $conn->query($query);
                $row2 = $result2->fetch_assoc();
                echo "<td colspan=2>" . $row2['total_count'] . "</td>";
                $totalKanan += $row2['total_count']; // Add the current row's total_count to totalKanan
            }
            echo "<td>" . $totalKanan . "</td>";
            ?>
        </tr>
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
        XLSX.writeFile(wb, 'LP Usia Korban ' + namefiletime + '.xlsx');
    }

    document.getElementById('exportButton').addEventListener('click', exportToExcel);
</script>
</html>

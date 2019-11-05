<?php

date_default_timezone_set('Asia/Jakarta');
ini_set('memory_limit', '1024M');
ini_set('max_execution_time', 0);
error_reporting(E_ALL);
ini_set('display_errors', 1);
header("Access-Control-Allow-Origin: *");

// delete wget file
foreach (glob("../sync-sqlsrv.php@type=*") as $filename) {
    unlink($filename);
}

$srvName  = "server\\sqlexpress, 49170"; //srvName\instanceName, portNumber (default is 1433)
$connInfo = array("Database" => "Anviz", "UID" => "absensi", "PWD" => "absensi");
$sqlsrv   = sqlsrv_connect($srvName, $connInfo);
$mysql    = new PDO("mysql:host=localhost;dbname=sgedb", 'root', 's3k4w4n');

if (!$sqlsrv) {
    echo "<h1>Connection could not be established (SQL SERVER) &#128546;</h1>";
    die(print_r(sqlsrv_errors(), true));
}

if (!$mysql) {
    echo "<h1>Connection could not be established (MySQL) &#128546;</h1>";
    die(print_r(mysql_errno(), true));
}

$type = isset($_GET['type']) || isset($_POST['type']);
if (!$type) {
    die('<h1>What do you want? add `type` query! &#128547;</h1>');
}

switch ($type) {
    case 'save_absensi_manual':
        saveAbsensiManual();
        break;
}

// Save absensi manual
function saveAbsensiManual()
{
    global $sqlsrv;
    global $mysql;
    $post = $_POST['data'];
    $period = [];
    $arrDt = [];
    if ($post['date'] !== '') {
        $dates = explode('|', $post['date']);
        $period = new DatePeriod(
            new DateTime($dates[0]),
            new DateInterval('P1D'),
            new DateTime($dates[1])
        );

        foreach ($period as $key => $dt) {
            array_push($arrDt, $dt->format('Y-m-d'));
        }
        if (!in_array($dates[1], $arrDt, true)) {
            array_push($arrDt, $dates[1]);
        }
    }
    $save = [];
    $save['nik'] = $post['nik'];
    $save['date'] = $post['dateSingle'];
    $save['keterangan'] = $post['keterangan'];
    $save['not_checkin'] = $post['reason']['not_checkin'];
    $save['not_checkout'] = $post['reason']['not_checkout'];
    $save['absent'] = $post['reason']['absent'];
    // SQL Server identity
    $id_on = "SET IDENTITY_INSERT Checkinout ON; ";
    $id_off = "SET IDENTITY_INSERT Checkinout OFF; ";
    // data to save
    $data = [
        'Userid' => $save['nik'],
        'CheckTime' => '',
        'CheckType' => '',
        'Sensorid' => 1,
        'Logid' => '',
        'Checked' => '',
        'WorkType' => 0,
        'AttFlag' => 1
    ];
    // sql server insert
    $ins_sqlsrv = "INSERT INTO checkinout 
            (Userid, CheckTime, CheckType, Sensorid, Logid, Checked, WorkType, AttFlag) 
            VALUES ";

    if ($save['not_checkin']) {

        $data['CheckType'] = 'I';
        $data['CheckTime'] = $save['date'] . ' 08:00:00';
        if (count($arrDt) > 0) {
            foreach ($arrDt as $dt) {
                $data['CheckTime'] = $dt;
                $sql = $ins_sqlsrv . " ('" . $save['nik'] . "', '" . $dt . " 08:00:00', 'I', 1, '', '', 0, 1); ";

                // Cek data in sql server
                $cek_sql = "SELECT * FROM checkinout WHERE Userid = '" . $save['nik'] . "' AND CheckTime = '" . $dt . " 08:00:00' ";
                $cek = sqlsrv_query($sqlsrv, $cek_sql);
                if (sqlsrv_num_rows($cek) === false) {
                    $insqlsrv = sqlsrv_query($sqlsrv, $id_on . $sql . $id_off);
                    $keys = implode(',', array_keys($data));
                    $values = implode("','", array_values($data));
                    $repmysql = " REPLACE INTO checkinout (" . $keys . ") VALUES ('" . $values . "')";
                    $stmt = $mysql->prepare($repmysql);
                    $stmt->execute();
                }
            }
        } else {
            $sql = $ins_sqlsrv . " ('" . $save['nik'] . "', '" . $save['date'] . " 08:00:00', 'I', 1, '', '', 0, 1); ";
            $cek_sql = "SELECT * FROM checkinout WHERE Userid = '" . $save['nik'] . "' AND CheckTime = '" . $save['date'] . " 08:00:00' ";
            $cek = sqlsrv_query($sqlsrv, $cek_sql);
            if (sqlsrv_num_rows($cek) === false) {
                $insqlsrv = sqlsrv_query($sqlsrv, $id_on . $sql . $id_off);
                $keys = implode(',', array_keys($data));
                $values = implode("','", array_values($data));
                $repmysql = " REPLACE INTO checkinout (" . $keys . ") VALUES ('" . $values . "')";
                $stmt = $mysql->prepare($repmysql);
                $stmt->execute();
            }
        }
    }

    if ($save['not_checkout']) {

        $data['CheckType'] = 'O';
        $data['CheckTime'] = $save['date'] . ' 17:00:00';
        if (count($arrDt) > 0) {
            foreach ($arrDt as $dt) {
                $data['CheckTime'] = $dt;
                $sql = $ins_sqlsrv . " ('" . $save['nik'] . "', '" . $dt . " 17:00:00', 'O', 1, '', '', 0, 1); ";

                // Cek data in sql server
                $cek_sql = "SELECT * FROM checkinout WHERE Userid = '" . $save['nik'] . "' AND CheckTime = '" . $dt . " 17:00:00' ";
                $cek = sqlsrv_query($sqlsrv, $cek_sql);
                if (sqlsrv_num_rows($cek) === false) {
                    $insqlsrv = sqlsrv_query($sqlsrv, $id_on . $sql . $id_off);
                    $keys = implode(',', array_keys($data));
                    $values = implode("','", array_values($data));
                    $repmysql = " REPLACE INTO checkinout (" . $keys . ") VALUES ('" . $values . "')";
                    $stmt = $mysql->prepare($repmysql);
                    $stmt->execute();
                }
            }
        } else {
            $sql = $ins_sqlsrv . " ('" . $save['nik'] . "', '" . $save['date'] . " 17:00:00', 'O', 1, '', '', 0, 1); ";
            $cek_sql = "SELECT * FROM checkinout WHERE Userid = '" . $save['nik'] . "' AND CheckTime = '" . $save['date'] . " 17:00:00' ";
            $cek = sqlsrv_query($sqlsrv, $cek_sql);
            if (sqlsrv_num_rows($cek) === false) {
                $insqlsrv = sqlsrv_query($sqlsrv, $id_on . $sql . $id_off);
                $keys = implode(',', array_keys($data));
                $values = implode("','", array_values($data));
                $repmysql = " REPLACE INTO checkinout (" . $keys . ") VALUES ('" . $values . "')";
                $stmt = $mysql->prepare($repmysql);
                $stmt->execute();
            }
        }
    }


    if (count($arrDt) > 0) {
        foreach ($arrDt as $dt) {
            $save['date'] = $dt;
            $keys = implode(',', array_keys($save));
            $values = implode("','", array_values($save));
            $repmysql = " REPLACE INTO tketabsent (" . $keys . ") VALUES ('" . $values . "')";
            $stmt = $mysql->prepare($repmysql);
            $stmt->execute();
        }
    } else {
        $keys = implode(',', array_keys($save));
        $values = implode("','", array_values($save));
        $repmysql = " REPLACE INTO tketabsent (" . $keys . ") VALUES ('" . $values . "')";
        $stmt = $mysql->prepare($repmysql);
        $stmt->execute();
    }

    header("Access-Control-Allow-Origin: *");
    echo json_encode(['data' => ['code' => 1, 'message' => '']]);
}

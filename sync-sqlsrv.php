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
$connInfo = array("Database" => "FP", "UID" => "sa", "PWD" => "s3k4w4n");
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

$type = isset($_GET['type']);
if (!$type) {
    die('<h1>What do you want? add `type` query! &#128547;</h1>');
}

switch ($type) {
    case 'sync_absensi':
        getAbsensiFromFingerPrint();
        break;
}

function getAbsensiFromFingerPrint()
{
    global $sqlsrv;
    global $mysql;
    $days = 0;
    $sql  = "SELECT days FROM `msettingtable` where `name` = 'sync_absensi_fingerprint' ";
    $stmt = $mysql->prepare($sql);
    $stmt->execute();
    $result = $stmt->fetch(PDO::FETCH_ASSOC);
    if ($stmt->rowCount() > 0) {
        $days = abs($result['days']);
    }

    $today = date_create(date('Y-m-d'));
    $dt    = date_format(date_sub($today, date_interval_create_from_date_string("$days days")), 'Y-m-d H:i:s');
    $sql   = "SELECT Userid,CheckTime,CheckType,Sensorid,Logid,Checked,WorkType,AttFlag FROM CheckInOut where CheckTime >= '$dt'";
    $stmt  = sqlsrv_query($sqlsrv, $sql);
    if ($stmt === false) {
        die(print_r(sqlsrv_errors(), true));
    }

    $i    = 0;
    $records = [];
    while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
        $row['CheckTime'] = date_format($row['CheckTime'], 'Y-m-d H:i:s');
        $records[]           = $row;

    }
    $col_arr = array_keys($records[0]);
    $col     = implode('`,`', $col_arr);
    $sql     = "REPLACE INTO `CheckInOut` (`" . $col . "`) VALUES ";
    $values = "";
    foreach ($records as $key => $val) {
        $val['Checked'] = 0;
        $val_arr = array_values($val);
        $val     = implode("','", $val_arr);
        $values .= "('" . $val . "'),";
    }
    $values = substr($values, 0, strlen($values) - 1);
    $sql  = $sql.$values;
    $stmt = $mysql->prepare($sql);
    $mysql->beginTransaction();
    try {
        $stmt = $mysql->prepare($sql);
        $stmt->execute();
        $mysql->commit();
        $msg = 'Sinkronisasi Sukses!';
        $success = TRUE;
        writeLog('Sukses!');
    } catch (Exception $e) {
        $mysql->rollBack();
        $msg = $e;
        $success = FALSE;
        writeLog($e);
    }
    
    echo json_encode(['success' => $success, 'message' => $msg]);
}

function writeLog($message)
{ 
    $message = date('Y-m-d H:i:s') . ' ' . $message . "\n";
    file_put_contents('./app-log.txt', $message, FILE_APPEND);
}

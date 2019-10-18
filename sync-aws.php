<?php
/**
 * created by @abdmun8 2019
 */
date_default_timezone_set('Asia/Jakarta');
ini_set('memory_limit', '1024M');
ini_set('max_execution_time', 0);
error_reporting(E_ALL);
ini_set('display_errors', 1);

// delete wget file
foreach (glob("../sync-aws.php?@type=sync_*") as $filename) {
    unlink($filename);
}

require 'vendor/autoload.php';

$mysql = new PDO("mysql:host=192.168.3.38;dbname=sgedb", 'root', 's3k4w4n');
// Check Request type
$type = isset($_GET['type']);
if (!$type) {
    die('<h1>What do you want? add `type` query! &#128547;</h1>');
}
// check connection to db status
if (!$mysql) {
    echo "<h1>Connection could not be established (MySQL) &#128546;</h1>";
    die(print_r(mysql_errno(), true));
}

switch ($type) {
    case 'sync_absensi':
        getAbsensiFromServer();
        break;
}


function getAbsensiFromServer()
{
    global $mysql;
    $sql = "SELECT days FROM `msettingtable` WHERE name ='sync_absensi_subcont' ";
    $stmt = $mysql->prepare($sql);
    $stmt->execute();
    $result = $stmt->fetch(PDO::FETCH_ASSOC);
    $days = abs($result['days']);
    $today = date_create(date('Y-m-d'));
    $dt = date_format(date_sub($today, date_interval_create_from_date_string("$days days")), 'Y-m-d');
    $url = 'http://api.sekawanpm.com/sync';
    $client = new GuzzleHttp\Client();
    $param = [
        'query' => [
            'limit_date' => $dt,
            'key' => '5b8e847c500b3a026824738c40246e75'
        ]
    ];
    $response = $client->request('GET', $url, $param);
    $result = json_decode($response->getBody()->getContents(), true);
    if($result['status'] != 'success' || count($result['data']) == 0){
        writeLog('Problem connect to server or data Null');
        die;
    }
    $field = implode(",",array_keys($result['data'][0]));
    $sql = " REPLACE INTO `checkinoutsubcont` ($field) VALUES ";
    $values = "";
    foreach ($result['data'] as $key => $v) {
        $values .= "('".implode("','",$v)."'),";
    }
    $nvalues = substr($values, 0, strlen($values) - 1);
    $sql  = $sql.$nvalues;
    $stmt = $mysql->prepare($sql);
    $mysql->beginTransaction();
    try {
        $stmt = $mysql->prepare($sql);
        $stmt->execute();
        $mysql->commit();
        writeLog('Sukses!');
    } catch (Exception $e) {
        $mysql->rollBack();
        writeLog($e);
    }
}

function writeLog($message)
{    
    $message = date('Y-m-d H:i:s') . ' ' . $message . "\n";
    file_put_contents('./sync-log-aws.txt', $message, FILE_APPEND);
}

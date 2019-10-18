<?php

/**
 * if you want to debug code, uncomment line below and add debug param in your request 
 */
// if(!isset($_GET['debug'])) die;

date_default_timezone_set('Asia/Jakarta');
ini_set('memory_limit', '1024M');
ini_set('max_execution_time', 0);
error_reporting(E_ALL);
ini_set('display_errors', 1);
// delete wget file
foreach (glob("../sync-mysql.ph*") as $filename) {
    unlink($filename);
}

require 'vendor/autoload.php';

use GuzzleHttp\Client;

$log        = '';
$log_data   = [];
$fail       = 0;
/* Send log to email */
$log_email = 'info@sekawanpm.com';
/* Connection */
$dbh = new PDO("mysql:host=server;dbname=sync_db", 'root', 's3k4w4n');
/* Table sync name */
$table_sync_name = "sync_db.table_to_sync_mysql";
/* SQL from table sync */
$sql_sync  = "SELECT * FROM {$table_sync_name} where `sync` = 1 ";
$table_sync = getDataFromTable($sql_sync);
$msg        = '[import started]';
writeLog($msg);
$log .= date('Y-m-d H:i:s') . '.' . explode('.', microtime(true))[1] . ' ' . $msg . "\n";
$n = 0;
$fail = 0;

// print_r($table_sync);
// die;

foreach ($table_sync as $key => $val) {
    $n++;
    $table_name_from = $val['table_name_from']; // table from sync
    $table_name_to = $val['table_name_to']; // table to sync
    $get_query = $val['get_query']; // insert query
    $delete_query = $val['delete_query']; //query delete
    writeLog($n . '. ' . $table_name_from . '->' . $table_name_to);

    /* Get mysql table structure */
    $data_table_from = getMysqlTable($table_name_from);
    $data_table_to = getMysqlTable($table_name_to);
    if ($data_table_from['success'] !== true && $data_table_to['success'] !== true) {
        $msg = 'failed fetch mysql table ' . $table_name_from . ' -> ' . $data_table_from['msg'] . ' || ' . $table_name_to . ' -> ' . $data_table_to['msg'];
        $log .= date('Y-m-d H:i:s') . '.' . explode('.', microtime(true))[1] . ' ' . $msg . "\n";
        writeLog($msg);
        $fail++;
        continue;
    }

    /* Compare mysql and dbf table column */
    $same = array_diff($data_table_from['columns'], $data_table_to['columns']);
    if (count($same) > 0) {
        $arr  = array_values($same);
        $diff = implode(',', $arr);
        $msg  = 'kolom table from dan to (' . $diff . ')';
        writeLog($msg);
        $log .= date('Y-m-d H:i:s') . '.' . explode('.', microtime(true))[1] . ' ' . $msg . "\n";
        $fail++;
        continue;
    }

    /* Delete data from table_to */
    $delete_record_table_to = deleteQuery($table_name_to, $delete_query);
    if ($delete_record_table_to !== TRUE) {
        $msg = 'failed delete data from ' . $table_name_to;
        $log .= date('Y-m-d H:i:s') . '.' . explode('.', microtime(true))[1] . ' ' . $msg . "\n";
        writeLog($msg);
        $fail++;
        continue;
    }
    
    
    /* Get data from table from */
    $data_from = getDataFromTable($get_query);    
    if (count($data_from) > 0) {
        $insert = insertToMysqlTable($data_from, $table_name_to);
        if ($insert !== true) {
            $msg = 'failed insert from ' . $table_name_from . ' -> ' . $table_name_to;
            $log .= date('Y-m-d H:i:s') . '.' . explode('.', microtime(true))[1] . ' ' . $msg . "\n";
            writeLog($msg);
            $fail++;
            continue;
        }
    }

    $msg = "import " . $table_name_from . ' -> ' . $table_name_to . " success!";
    writeLog($msg);
    $log .= date('Y-m-d H:i:s') . '.' . explode('.', microtime(true))[1] . ' ' . $msg . "\n";
}

$msg = '[import finished]';
writeLog($msg);

// die;
if ($fail > 0) {
    $client   = new Client();
    $response = $client->request('POST', 'https://apps.sekawanpm.com/mail.php', [
        'form_params' => [
            'token'   => '9ab6578fc863f7ea13cf108bd6c6e499',
            'to'      => $log_email,
            'subject' => 'Error Log Import Mysql to Mysql',
            'msg'     => $log,
        ],
    ]);
}

function writeLog($message)
{
    global $log_data;
    array_push($log_data, [date('Y-m-d H:i:s') . '.' . explode('.', microtime(true))[1], $message]);
    $message = date('Y-m-d H:i:s') . '.' . explode('.', microtime(true))[1] . ' ' . $message . "\n";
    file_put_contents('./sync-mysql-log.txt', $message, FILE_APPEND);
}

/* Get data from table */
function getDataFromTable($sql)
{
    global $dbh;
    $data = [];
    $stmt = $dbh->prepare($sql);
    $stmt->execute();
    $result = $stmt->fetchAll(PDO::FETCH_ASSOC);
    if ($stmt->rowCount() > 0) {
        return $result;
    }

    return $data;
}

// delete and insert data
function deleteQuery($table_name_to, $sql)
{
    global $dbh;
    $dbh->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $dbh->beginTransaction();
    try {
        $stmt = $dbh->prepare($sql);
        $stmt->execute();
        $dbh->commit();
        return TRUE;
    } catch (Exception $e) {
        $dbh->rollBack();
        return 'failed delete table ' . $table_name_to . ' > ' . $e->getMessage();
    }
}

// get data from mysql table
function getDataMysql($tbl_nm)
{
    global $dbh;

    $sql  = "SELECT * FROM " . $tbl_nm . " ";
    $stmt = $dbh->prepare($sql);
    $stmt->execute();
    $res = $stmt->fetchAll(PDO::FETCH_ASSOC);
    if ($stmt->rowCount() > 0) {
        return $res;
    }

    return [];
}


// insert table mysql
function insertToMysqlTable($records, $tbl_nm)
{
    global $dbh;
    $col_arr = array_keys($records[0]);
    $col     = implode('`,`', $col_arr);
    $sql     = "INSERT INTO " . $tbl_nm . " (`" . $col . "`) VALUES ";
    $values  = '';
    try {
        $dbh->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        $dbh->beginTransaction();
        foreach ($records as $key => $val) {
            $val_arr = array_values($val);
            $val     = implode("','", $val_arr);
            $values .= "('" . $val . "'),";     

        }
        $values = substr($values, 0, strlen($values) - 1);
        $stmt   = $dbh->prepare($sql . $values);
        $stmt->execute();
        $dbh->commit();
        return TRUE;
    } catch (Exception $e) {
        $dbh->rollBack();
        return $e->getMessage();
    }
}

// get data and column mysql
function getMysqlTable($tbl_nm)
{
    global $dbh;
    $data = [
        'success' => false,
        'columns' => [],
        'msg'     => 'Table ' . $tbl_nm . ' belum ada pada database MySQL!',
    ];

    $sql  = "SHOW COLUMNS FROM " . $tbl_nm . " ";
    $stmt = $dbh->prepare($sql);
    $stmt->execute();
    $result = $stmt->fetchAll(PDO::FETCH_ASSOC);
    if ($stmt->rowCount() > 0) {
        $data = [
            'success' => true,
            'columns' => getMysqlColumn($result),
            'msg'     => 'ok',
        ];
    }
    return $data;
}

// get mysql table column
function getMysqlColumn($result)
{
    $no   = 0;
    $data = [];
    foreach ($result as $key => $value) {
        array_push($data, $value['Field']);
    }
    return $data;
}

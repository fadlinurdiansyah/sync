<?php

date_default_timezone_set('Asia/Jakarta');
ini_set('memory_limit', '1024M');
ini_set('max_execution_time', 0);
error_reporting(E_ALL);
ini_set('display_errors', 1);

// delete wget file
foreach (glob("../sync-v1.php@create=*") as $filename) {
    unlink($filename);
}

/**
 * you should create select query for the new table there
 * status log saved to table_sync_status
 * created by @abdmun8 2019
 */

$con = new PDO("mysql:host=localhost;dbname=sync_db", 'root', 's3k4w4n');
$start_time = '';

if (isset($_GET['create'])) {

    $sqlget = "SELECT * FROM table_sync_create WHERE active = 1  ";
    $stmt = $con->prepare($sqlget);
    $stmt->execute();
    $res = $stmt->fetchAll(PDO::FETCH_OBJ);
    if (count($res) > 0) {
        foreach ($res as $key => $row) {
            $GLOBALS['table_name'] = $row->table_name;
            $GLOBALS['sql'] = $row->sql_select;
            $mysqlcon = new PDO("mysql:host=localhost;dbname={$row->mysqldb}", 'root', 's3k4w4n');
            createTableYouWant($mysqlcon);
        }
    }
} else {
    echo 'Welcome to sync Home :)';
}

/* Method should write Below */

function createTableYouWant($mysqlcon)
{    
    global $start_time;
    $start_time = date('Y-m-d H:i:s');
    $mysqlcon->beginTransaction();
    // create table temp 
    $sqlTmp = "CREATE TEMPORARY TABLE temp_" . $GLOBALS['table_name'] . " " . $GLOBALS['sql'] . " ";
    try {
        $stmt = $mysqlcon->prepare($sqlTmp);
        $stmt->execute();
    } catch (Exception $e) {
        $mysqlcon->rollBack();
        saveLogTable(0, $e);
        return FALSE;
    }
    // delete real table
    $start_time = date('Y-m-d H:i:s');
    $sqlDel = "DROP TABLE IF EXISTS " . $GLOBALS['table_name'] . " ";
    try {
        $stmt = $mysqlcon->prepare($sqlDel);
        $stmt->execute();
    } catch (Exception $e) {
        $mysqlcon->rollBack();
        saveLogTable(0, $e);
        return FALSE;
    }
    // create real table
    $sqlCrt = "CREATE TABLE " . $GLOBALS['table_name'] . " SELECT * FROM temp_" . $GLOBALS['table_name'] . " ";
    try {
        $stmt = $mysqlcon->prepare($sqlCrt);
        $stmt->execute();
        $mysqlcon->commit();
        saveLogTable();
    } catch (Exception $e) {
        $mysqlcon->rollBack();
        saveLogTable(0, $e);
        return FALSE;
    }
    return TRUE;
}

function saveLogTable($success = 1, $msg = 'Success')
{
    global $con;
    global $start_time;
    $sql = "REPLACE INTO table_sync_status (`table_name`,`start_time`,`success`,`message`) VALUES (?,?,?,?)";
    $stmt = $con->prepare($sql);
    $stmt->execute([$GLOBALS['table_name'], $start_time, $success, $msg]);
}

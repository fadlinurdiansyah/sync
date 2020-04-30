<?php
//Test Commit
date_default_timezone_set('Asia/Jakarta');
ini_set('memory_limit', '1024M');
ini_set('max_execution_time', 0);
error_reporting(E_ALL);
ini_set('display_errors', 1);

require_once 'vendor/autoload.php';

use GuzzleHttp\Client;

// delete wget file
foreach (glob("../auto_mailer.php*") as $filename) {
    unlink($filename);
}

/**
 * you should create select query for the new table there
 * status log saved to table_sync_status
 * created by @abdmun8 2019
 */

$consgedb = new PDO("mysql:host=server;dbname=sgedb", 'root', 's3k4w4n');
$consyncdb = new PDO("mysql:host=server;dbname=sync_db", 'root', 's3k4w4n');

$type = isset($_GET['type']) ? TRUE : FALSE;
if (!$type) {
    die('<h1>What do you want? add `type` query! &#128547;</h1>');
}

/** 
 * Send Mail Menu:
 * 1. Send Mail SIC : type = send_mail_sic
 */

switch ($_GET['type']) {
    case 'send_mail_sic':
        sendMailSIC();
        break;
    case 'send_mail_wo_over_due':
        sendMailWOOverDue();
        break;
    case 'send_mail_po_price_more_than_std':
        sendMailPOPriceMoreTanStd();
        break;
}

function sendMailSIC()
{
    global $consgedb;
    global $consyncdb;

    /* Get data Report sgedb */
    $sql = " SELECT 
                    stcd,
                    nama,
                    spek,
                    maker,
                    uom,
                    buffer,
                    reorder,
                    bal_stock,
                    prouts,
                    poouts,
                    IF(IF(bal_stock <= reorder,
                            buffer - prouts - poouts,
                            0) < 0,
                        0,
                        IF(bal_stock <= reorder,
                            buffer - prouts - poouts,
                            0)) AS advice
                FROM
                    (SELECT 
                        mstchd.stcd,
                            mstchd.nama,
                            mstchd.spek,
                            mstchd.maker,
                            mstchd.uom,
                            mstchd.buffer,
                            mstchd.reorder,
                            IFNULL(mstcdt.begbal, 0) AS begbal,
                            IFNULL(mstcdt.receive, 0) AS recive,
                            IFNULL(mstcdt.issue, 0) AS issue,
                            IFNULL(((mstcdt.begbal + mstcdt.receive) - mstcdt.issue), 0) AS bal_stock,
                            COALESCE(prpo_outs.outspo, 0) AS prouts,
                            COALESCE(prpo_outs.outsrecive, 0) AS poouts
                    FROM
                        mstchd
                    LEFT JOIN (SELECT 
                        stcd,wono,sum(begbal) as begbal, sum(issue) as issue, sum(receive) as receive
                    FROM
                        mstcdt
                    WHERE
                        wono IN ('INV' , 'CNS') group by stcd) AS mstcdt ON mstchd.stcd = mstcdt.stcd
                    LEFT JOIN prpo_outs ON prpo_outs.stcd = mstcdt.stcd
                    WHERE
                        mstchd.buffer > 0) AS SIC
                WHERE
                    (sic.bal_stock + sic.prouts + sic.poouts) < sic.buffer
                        AND stcd LIKE 'CNS%'
                HAVING advice > 0";
    $stmt = $consgedb->prepare($sql);
    $stmt->execute();
    $data = $stmt->fetchAll(PDO::FETCH_ASSOC);

    /* Data Mail from sync_db.send_mail */
    $sql_get = " SELECT * FROM send_mail WHERE name = '" . $_GET['type'] . "' ";
    $stmt = $consyncdb->prepare($sql_get);
    $stmt->execute();

    /* Title Document */
    $title_document = "temp/Report_SIC_Stock_" . date('dmY') . ".xlsx";
    require_once 'views/report_sic.php';

    $data_mail = $stmt->fetch(PDO::FETCH_ASSOC);
    $multipart = [
        [
            'name'     => 'token',
            'contents' => '9ab6578fc863f7ea13cf108bd6c6e499'
        ],
        [
            'name'     => 'subject',
            'contents' => $data_mail['subject'] . ' ' . date('d F Y')
        ],
        [
            'name'     => 'msg',
            'contents' => $data_mail['msg']
        ],
        [
            'name'     => 'file',
            'contents' => fopen($title_document, 'r')
        ]

    ];

    $params = stringToArrayParam($multipart, ['to' => $data_mail['to'], 'cc' => $data_mail['cc']]);

    /* Send Request */
    $client = new Client();
    $response = $client->request('POST', 'http://192.168.3.21/sekawan-site-menu/mail.php', [
        'multipart' => $params
    ]);

    echo $response->getBody()->getContents();
}

function stringToArrayParam($multipart = [], $data)
{
    $params = $multipart;
    foreach ($data as $key => $value) {
        $arr = explode(',', $value);
        for ($i = 0; $i < count($arr); $i++) {
            $temp = [
                'name' => $key . '[' . $i . ']',
                'contents' => trim($arr[$i])
            ];
            array_push($params, $temp);
        }
    }
    return $params;
}

function sendMailWOOverDue()
{
    global $consgedb;
    global $consyncdb;

    $dayNo = date('w');
    $sql_having = "";
    $sql_having = " HAVING `left` <= 14 AND date_close = 'N' ";
    // if($dayNo == 1){
    //     $sql_having = " HAVING `left` = 14 OR `left` = 13 OR `left` = 12 OR `left` <= 0 ";
    // }else{
    //     $sql_having = " HAVING `left` = 14 OR `left` <= 0 ";
    // }

    /* Get data Report sgedb */
    $sqlWo = " SELECT a.wono,a.`desc`,a.`date`,a.selesai,(CASE WHEN h.approve_mkt IS NULL OR h.approve_mkt='' THEN DATEDIFF(a.selesai, CURDATE()) WHEN h.approve_mkt='Y' THEN DATEDIFF(a.selesai,h.date_close) END) AS `left`,a.complete,a.jenis,a.`type`,a.category,e.nama,b.tseng,c.tsmfc,d.item_total, CONCAT(SUBSTR(f.nama,1,7),' / ', SUBSTR(k.nama,1,7)) AS mkt, CONCAT(SUBSTR(g.nama,1,7),' / ', SUBSTR(j.nama,1,7)) AS pl, IFNULL(h.date_close,'N') AS date_close, IFNULL(h.approve_mkt,'N') AS approve_close,i.conter, IFNULL(i.approve_extend,'N') AS approve_extend
    FROM trwo a
    LEFT JOIN ts_eng2 b ON b.wono=a.wono
    LEFT JOIN ts_mfc2 c ON c.wono=a.wono
    LEFT JOIN pr_total d ON d.wono=a.wono
    LEFT JOIN customer e ON a.custid=e.custid
    LEFT JOIN personal AS f ON a.pic=f.id_personalia
    LEFT JOIN personal AS g ON a.pl=g.id_personalia
    LEFT JOIN close_wo_pl h ON a.wono=h.wono
    LEFT JOIN (
    SELECT wono, COUNT(wono) AS conter,approv_mkt AS approve_extend
    FROM extend_wo
    GROUP BY wono) AS i ON i.wono=a.wono
    LEFT JOIN personal AS j ON g.atasan=j.id_personalia
    LEFT JOIN personal AS k ON f.atasan=k.id_personalia
    WHERE a.complete='1899-12-30' AND a.jenis IN(1,2,3,4,5)
    $sql_having
    ORDER BY `left` DESC ";
    $stmt = $consgedb->prepare($sqlWo);
    $stmt->execute();
    $data = $stmt->fetchAll(PDO::FETCH_ASSOC);

    /* Data Mail from sync_db.send_mail */
    $sql_get = " SELECT * FROM send_mail WHERE name = '" . $_GET['type'] . "' ";
    $stmt = $consyncdb->prepare($sql_get);
    $stmt->execute();

    /* Title Document */
    $title_document = "temp/Report_WO_Overdue_" . date('dmY') . ".xlsx";
    require_once 'views/report_wo_overdue.php';

    $data_mail = $stmt->fetch(PDO::FETCH_ASSOC);
    $multipart = [
        [
            'name'     => 'token',
            'contents' => '9ab6578fc863f7ea13cf108bd6c6e499'
        ],
        [
            'name'     => 'subject',
            'contents' => $data_mail['subject'] . ' ' . date('d F Y')
        ],
        [
            'name'     => 'msg',
            'contents' => $data_mail['msg']
        ],
        [
            'name'     => 'file',
            'contents' => fopen($title_document, 'r')
        ]

    ];

    $params = stringToArrayParam($multipart, ['to' => $data_mail['to'], 'cc' => $data_mail['cc']]);

    /* Send Request */
    $url = 'http://192.168.3.21/sekawan-site-menu/mail.php';
    // $url = '192.168.3.224/sekawan-site-menu/mail.php';
    $client = new Client();
    $response = $client->request('POST', $url, [
        'multipart' => $params
    ]);

    echo $response->getBody()->getContents();
}

// Send mail Report Harga PO Lebih dari Harga Standar
function sendMailPOPriceMoreTanStd()
{
    global $consgedb;
    global $consyncdb;

    /* Get data Report sgedb */
    $sqlWo = " SELECT * from (SELECT 
            a.wono,
            b.date,
            b.pono,
            vendor.nama,
            a.stcd,
            a.desc,
            a.qtyodr,
            (CASE
                WHEN b.curr = 'IDR' THEN (a.gross - (a.gross * (b.disc / b.jumlah)))
                WHEN b.curr <> 'IDR' THEN ((a.gross - (a.gross * (b.disc / b.jumlah))) * f.tkrsac)
            END) AS price,
            ifnull(msprice.pur,0) as stdprice,
            msprice.remark
        FROM
            podt a
                LEFT JOIN
            pohd b ON a.pono = b.pono
                LEFT JOIN
            tbkurs f ON CONCAT(f.tkrskd, f.tkrsdk) = CONCAT(b.curr,
                    YEAR(b.`date`),
                    MID(b.`date`, 6, 2))
                LEFT JOIN
            msprice ON a.stcd = msprice.stcd
                LEFT JOIN
            vendor ON b.venid = vendor.venid
        WHERE
            a.status <> 'C'
                AND b.date = CURDATE() - 1)
                -- '2019-10-18') 
                as compare where stdprice <> 0 and price > stdprice ";
    $stmt = $consgedb->prepare($sqlWo);
    $stmt->execute();
    $result = $stmt->fetchAll(PDO::FETCH_ASSOC);

    /* die if theres no record */
    if (count($result) == 0) die;

    // $data = [];
    $no = 0;
    $date = date('d/m/Y', strtotime($result[0]['date']));
    $content = "<h2>PT. SEKAWAN GLOBAL ENGINEERING</h2>";
    $content .= "<h3>Report Harga PO Lebih dari Harga Standar</h3>";
    $content .= "<h3>Periode " . $date . "</h3>";
    $content .= "<table border='1' cellspacing='0' cellpadding='3'>";
    $content .= "<tr>
            <th>NO</th>
            <th>PO NO</th>
            <th>DATE</th>
            <th>SUPPLIER</th>
            <th>WONO</th>
            <th>ITEM CODE</th>
            <th>ITEM NAME</th>
            <th>QTY</th>
            <th>HARGA</th>
            <th>STD HARGA</th>
            <th>SELISIH HARGA</th>
            <th>SELISIH (%)</th>
            <th>NOTE STD PRICE</th>
        </tr>";
    foreach ($result as $key => $row) {
        $no++;
        $qty = floatval($row['qtyodr']);
        $std_price = floatval($row['stdprice']);
        $price = floatval($row['price']);
        $selisih = $price - $std_price;
        $percent = round($selisih / $row['stdprice'] * 100);
        $content .= "<tr>
            <td>{$no}</td>
            <td>{$row['pono']}</td>
            <td>{$date}</td>
            <td>{$row['nama']}</td>
            <td>{$row['wono']}</td>
            <td>{$row['stcd']}</td>
            <td>{$row['desc']}</td>
            <td style='text-align:right;'>{$qty}</td>
            <td style='text-align:right;'>" . number_format($price) . "</td>
            <td style='text-align:right;'>" . number_format($std_price) . "</td>
            <td style='text-align:right;'>" . number_format($selisih) . "</td>
            <td style='text-align:right;'>{$percent} %</td>
            <td>{$row['remark']}</td>
        </tr>";
    }
    $content .= "</table>";

    /* Data Mail from sync_db.send_mail */
    $sql_get = " SELECT * FROM send_mail WHERE name = '" . $_GET['type'] . "' ";
    $stmt = $consyncdb->prepare($sql_get);
    $stmt->execute();

    $data_mail = $stmt->fetch(PDO::FETCH_ASSOC);
    $multipart = [
        [
            'name'     => 'token',
            'contents' => '9ab6578fc863f7ea13cf108bd6c6e499'
        ],
        [
            'name'     => 'subject',
            'contents' => $data_mail['subject'] . ' ' . $date
        ],
        [
            'name'     => 'msg',
            'contents' => $content . "Pesan ini dikirim secara otomatis, jangan dibalas!"
        ],

    ];

    $params = stringToArrayParam($multipart, ['to' => $data_mail['to'], 'cc' => $data_mail['cc']]);

    /* Send Request */
    $url = 'http://192.168.3.21/sekawan-site-menu/mail.php';
    $client = new Client();
    $response = $client->request('POST', $url, [
        'multipart' => $params
    ]);

    echo $response->getBody()->getContents();
}

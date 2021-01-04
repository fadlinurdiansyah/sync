<?php

// var_dump($dataSummary);die;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

$H1 = [
    'NIK',
    'NAMA',
    'DEPT',
    'TS SPK',
    'NON SPK',
    'UNABS.',
    'TOTAL',
    '',
    'DAYS',
    'HOUR',
    'T.REGULAR',
    'OVERTIME',
    'G.TOTAL',
];


$H2 = [
    'TRX NO',
    'TANGGAL',
    'NIK',
    'NAMA',
    'DEPT',
    'WONO',
    'TS SPK',
    'NON SPK',
    'UNABS.',
    'ACTIVITY',
];

$arrHeaderAbsenTS = [
    'No.',
    'NIK',
    'Name',
    'Dept ID',
    'Date',
    'Duration'
];

$arrHeaderAbsenHoliday = [
    'No.',
    'NIK',
    'Name',
    'Dept ID',
    'Date',
    'Absensi'
];

$headerStyle = [
    'fill' => [
        'fillType' => Fill::FILL_GRADIENT_LINEAR,
        'rotation' => 90,
        'startColor' => [
            'argb' => 'ebf1de',
        ],
        'endColor' => [
            'argb' => 'ebf1de',
        ],
    ],
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => '00000000'],
        ],
    ],
    'font' => [
        'size' => 11,
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
    ],
];

$footerStyle = array_filter($headerStyle, function ($key) {
    return $key != 'alignment';
}, ARRAY_FILTER_USE_KEY);


$innerStyle = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => '00000000'],
        ],
    ],
    'font' => [
        'size' => 11
    ],
];

// Re arrange data
$dataSummary = array_values($dataSummary);

$spreadsheet = new Spreadsheet();
$myWorkSheet = new Worksheet($spreadsheet, "SUMMARY TS NON SPK");
$spreadsheet->addSheet($myWorkSheet, 0);
$spreadsheet->setActiveSheetIndex(0);
$activeSheet = $spreadsheet->getActiveSheet();
$activeSheet->setTitle("SUMMARY TS NON SPK");
$spreadsheet->getDefaultStyle()->getFont()->setName('Calibri');

// Set Title
$activeSheet->setCellValue('A1', 'PT. SEKAWAN GLOBAL ENGINEERING');
$activeSheet->setCellValue('A2', "REPORT SUMMARY TS NON SPK");
$activeSheet->setCellValue('A3', "PERIODE : " . $period);
$activeSheet->getStyle("A1:A3")->getFont()->setBold(true);

// Set Header
$activeSheet->fromArray($H1, NULL, "A6");
$activeSheet->getStyle("A6:G6")->applyFromArray($headerStyle);
$activeSheet->getStyle("I6:M6")->applyFromArray($headerStyle);
// $activeSheet->getStyle("A6:F6")->getFont()->setBold(true);
// $activeSheet->getStyle("A6:F6")->getFont()->setSize(11);

// set alignment
$activeSheet->getStyle("A6:F6")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
$activeSheet->getStyle("A6:F6")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$activeSheet->mergeCells("C5:G5");
$activeSheet->setCellValue('C5', 'ALLOCATION OF MAN-HOUR');
$activeSheet->getStyle("C5:G5")->applyFromArray(
    [
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'startColor' => ['argb' => '00000000'],
            ],
        ],
        'font'  => [
            'size' => 11,
            'name' => 'Calibri',
            'bold' => true,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
        ],
    ]
);


$activeSheet->mergeCells("I5:M5");
$activeSheet->setCellValue('I5', 'MAN-HOUR CAPACITY');
$activeSheet->getStyle("I5:M5")->applyFromArray(
    [
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'startColor' => ['argb' => '00000000'],
            ],
        ],
        'font'  => [
            'size' => 11,
            'name' => 'Calibri',
            'bold' => true,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
        ],
    ]
);


$row_start = 7;
$row = $row_start;
$data_dept = [];

// configure subtotal
$add = 7;
$no = 0;
$start = $add;
$end = 0;
$ts_spk = 0;
$non_spk = 0;
$unabs = 0;
$total = 0;
$totReg = 0;
$totOT = 0;
$totMHCap = 0;

for ($i = 0; $i < count($dataSummary); $i++) {
    $row = $i + $add;
    $activeSheet->setCellValue('A' . $row, $dataSummary[$i]['nik']);
    $activeSheet->setCellValue('B' . $row, $dataSummary[$i]['nama']);
    $activeSheet->setCellValue('C' . $row, $dataSummary[$i]['deptid']);
    $activeSheet->setCellValue('D' . $row, $dataSummary[$i]['ts_spk']);
    $activeSheet->setCellValue('E' . $row, $dataSummary[$i]['non_spk']);
    $activeSheet->setCellValue('F' . $row, $dataSummary[$i]['unabs']);
    $activeSheet->setCellValue('G' . $row, "=SUM(D$row:F$row)");
    $activeSheet->setCellValue('I' . $row, $dataSummary[$i]['days']);
    $activeSheet->setCellValue('J' . $row, $dataSummary[$i]['hours']);
    $activeSheet->setCellValue('K' . $row, "=(I$row*J$row)");
    $activeSheet->setCellValue('L' . $row, $dataSummary[$i]['ot']);
    $activeSheet->setCellValue('M' . $row, "=SUM(K$row:L$row)");

    $activeSheet->getRowDimension($row)->setOutlineLevel(1);
    $activeSheet->getStyle("A$row:G$row")->applyFromArray($innerStyle);
    $activeSheet->getStyle("I$row:M$row")->applyFromArray($innerStyle);

    if ($i == 0 && isset($dataSummary[$i + 1])) {
        if ($dataSummary[$i]['deptid'] != $dataSummary[$i + 1]['deptid']) {
            $end = $row;
            $totsub = $row + 1;
            $add += 1;
            $activeSheet->setCellValue("D$totsub", "=SUBTOTAL(9,D$start:D$end)");
            $activeSheet->setCellValue("E$totsub", "=SUBTOTAL(9,E$start:E$end)");
            $activeSheet->setCellValue("F$totsub", "=SUBTOTAL(9,F$start:F$end)");
            $activeSheet->setCellValue("G$totsub", "=SUBTOTAL(9,G$start:G$end)");
            $activeSheet->setCellValue("C$totsub", $dataSummary[$i]['deptid'] . " Total");
            $activeSheet->getStyle("A$totsub:G$totsub")->applyFromArray($footerStyle);

            $activeSheet->setCellValue("K$totsub", "=SUBTOTAL(9,K$start:K$end)");
            $activeSheet->setCellValue("L$totsub", "=SUBTOTAL(9,L$start:L$end)");
            $activeSheet->setCellValue("M$totsub", "=SUBTOTAL(9,M$start:M$end)");
            $activeSheet->setCellValue("I$totsub", $dataSummary[$i]['deptid'] . " Total");
            $activeSheet->getStyle("I$totsub:M$totsub")->applyFromArray($footerStyle);


            $start = $i + $add + 1;
            array_push($data_dept, $dataSummary[$i]['deptid']);
        }
    }

    if ($i != 0 && isset($dataSummary[$i + 1])) {
        if ($dataSummary[$i]['deptid'] != $dataSummary[$i + 1]['deptid']) {
            $end = $row;
            $totsub = $row + 1;
            $add += 1;
            $activeSheet->setCellValue("D$totsub", "=SUBTOTAL(9,D$start:D$end)");
            $activeSheet->setCellValue("E$totsub", "=SUBTOTAL(9,E$start:E$end)");
            $activeSheet->setCellValue("F$totsub", "=SUBTOTAL(9,F$start:F$end)");
            $activeSheet->setCellValue("G$totsub", "=SUBTOTAL(9,G$start:G$end)");
            $activeSheet->setCellValue("C$totsub", $dataSummary[$i]['deptid'] . " Total");
            $activeSheet->getStyle("A$totsub:G$totsub")->applyFromArray($footerStyle);

            $activeSheet->setCellValue("K$totsub", "=SUBTOTAL(9,K$start:K$end)");
            $activeSheet->setCellValue("L$totsub", "=SUBTOTAL(9,L$start:L$end)");
            $activeSheet->setCellValue("M$totsub", "=SUBTOTAL(9,M$start:M$end)");
            $activeSheet->setCellValue("I$totsub", $dataSummary[$i]['deptid'] . " Total");
            $activeSheet->getStyle("I$totsub:M$totsub")->applyFromArray($footerStyle);

            $start = $i + $add + 1;
            array_push($data_dept, $dataSummary[$i]['deptid']);
        }
    }

    if ($i == (count($dataSummary) - 1)) {
        $end = $row;
        $totsub = $row + 1;
        $add += 1;
        array_push($data_dept, $dataSummary[$i]['deptid']);
        $activeSheet->setCellValue("D$totsub", "=SUBTOTAL(9,D$start:D$end)");
        $activeSheet->setCellValue("E$totsub", "=SUBTOTAL(9,E$start:E$end)");
        $activeSheet->setCellValue("F$totsub", "=SUBTOTAL(9,F$start:F$end)");
        $activeSheet->setCellValue("G$totsub", "=SUBTOTAL(9,G$start:G$end)");
        $activeSheet->setCellValue("C$totsub", $dataSummary[$i]['deptid'] . " Total");
        $activeSheet->getStyle("A$totsub:G$totsub")->applyFromArray($footerStyle);

        $activeSheet->setCellValue("K$totsub", "=SUBTOTAL(9,K$start:K$end)");
        $activeSheet->setCellValue("L$totsub", "=SUBTOTAL(9,L$start:L$end)");
        $activeSheet->setCellValue("M$totsub", "=SUBTOTAL(9,M$start:M$end)");
        $activeSheet->setCellValue("I$totsub", $dataSummary[$i]['deptid'] . " Total");
        $activeSheet->getStyle("I$totsub:M$totsub")->applyFromArray($footerStyle);
    }
    $ts_spk += $dataSummary[$i]['ts_spk'];
    $non_spk += $dataSummary[$i]['non_spk'];
    $unabs += $dataSummary[$i]['unabs'];
    $total += ($dataSummary[$i]['ts_spk'] + $dataSummary[$i]['non_spk'] + $dataSummary[$i]['unabs']);

    $totReg += ($dataSummary[$i]['days'] * $dataSummary[$i]['hours']);
    $totOT += $dataSummary[$i]['ot'];
    $totMHCap += (($dataSummary[$i]['days'] * $dataSummary[$i]['hours']) + $dataSummary[$i]['ot']);
}
$rowGTotal = $row + 2;

$activeSheet->setCellValue('C' . $rowGTotal, 'GRAND TOTAL');
$activeSheet->setCellValue('D' . $rowGTotal, $ts_spk);
$activeSheet->setCellValue('E' . $rowGTotal, $non_spk);
$activeSheet->setCellValue('F' . $rowGTotal, $unabs);
$activeSheet->setCellValue('G' . $rowGTotal, $total);
$activeSheet->getStyle("A$rowGTotal:G$rowGTotal")->applyFromArray($footerStyle);

$activeSheet->setCellValue('I' . $rowGTotal, 'GRAND TOTAL');
$activeSheet->setCellValue('K' . $rowGTotal, $totReg);
$activeSheet->setCellValue('L' . $rowGTotal, $totOT);
$activeSheet->setCellValue('M' . $rowGTotal, $totMHCap);
$activeSheet->getStyle("I$rowGTotal:M$rowGTotal")->applyFromArray($footerStyle);

$rowTitleAbsenTS = $row + 4;
$activeSheet->setCellValue('A' . $rowTitleAbsenTS, "List Karyawan Belum Input TS");
$activeSheet->getStyle("A" . $rowTitleAbsenTS)->getFont()->setSize(14);
$activeSheet->getStyle("A" . $rowTitleAbsenTS)->getFont()->setBold(true);

$rowHeaderAbsen = $row + 5;
$row =  $rowHeaderAbsen + 1;
$no = 0;
foreach ($absTS as $key => $value) {
    $no++;
    $activeSheet->setCellValue('A' . $row, $no);
    $activeSheet->setCellValue('B' . $row, $value['nik']);
    $activeSheet->setCellValue('C' . $row, $value['nama']);
    $activeSheet->setCellValue('D' . $row, $value['departemen']);
    $activeSheet->setCellValue('E' . $row, $value['date']);
    $activeSheet->setCellValue('F' . $row, $value['TS']);
    $activeSheet->getStyle("A$row:F$row")->applyFromArray(
        [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'startColor' => ['argb' => '00000000'],
                ],
            ],
            'font'  => [
                'size' => 11,
                'name' => 'Calibri',
            ]
        ]
    );

    $row++;
}

$activeSheet->getStyle('A' . $rowHeaderAbsen . ':F' . $rowHeaderAbsen)->applyFromArray($headerStyle);
$activeSheet->fromArray($arrHeaderAbsenTS, NULL, 'A' . $rowHeaderAbsen);


$rowTitleAbsenHoliday = $row + 4;
$activeSheet->setCellValue('A' . $rowTitleAbsenHoliday, "List Karyawan Tidak Masuk/Off");
$activeSheet->getStyle("A" . $rowTitleAbsenHoliday)->getFont()->setSize(14);
$activeSheet->getStyle("A" . $rowTitleAbsenHoliday)->getFont()->setBold(true);

$rowHeaderAbsenHoliday = $row + 5;
$row =  $rowHeaderAbsenHoliday + 1;
$no = 0;
foreach ($absHoliday as $key => $value) {
    $no++;
    $activeSheet->setCellValue('A' . $row, $no);
    $activeSheet->setCellValue('B' . $row, $value['nik']);
    $activeSheet->setCellValue('C' . $row, $value['nama']);
    $activeSheet->setCellValue('D' . $row, $value['departemen']);
    $activeSheet->setCellValue('E' . $row, $value['date']);
    $activeSheet->setCellValue('F' . $row, $value['TS']);
    $activeSheet->getStyle("A$row:F$row")->applyFromArray(
        [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'startColor' => ['argb' => '00000000'],
                ],
            ],
            'font'  => [
                'size' => 11,
                'name' => 'Calibri',
            ]
        ]
    );

    $row++;
}

$activeSheet->getStyle('A' . $rowHeaderAbsenHoliday . ':F' . $rowHeaderAbsenHoliday)->applyFromArray($headerStyle);
$activeSheet->fromArray($arrHeaderAbsenHoliday, NULL, 'A' . $rowHeaderAbsenHoliday);


// define row size
$row_size = [
    'A' => 11,
    'B' => 25,
    'C' => 14,
    'D' => 10,
    'E' => 10,
    'F' => 10,
    'K' => 11,
    'L' => 11,
];

foreach ($row_size as $key => $value) {
    $row = explode(',', $key);
    foreach ($row as $col) {
        $activeSheet->getColumnDimension($col)->setWidth($value);
    }
}

/* Freeze celll */
$activeSheet->freezePane('A7');


// Create Detail TS multisheet
$n = 1;
foreach ($data_dept as $key => $d) {

    $myWorkSheet = new Worksheet($spreadsheet, $d);
    $spreadsheet->addSheet($myWorkSheet, $n);
    $spreadsheet->setActiveSheetIndex($n);
    $activeSheet = $spreadsheet->getActiveSheet();
    $activeSheet->setTitle($d);
    $spreadsheet->getDefaultStyle()->getFont()->setName('Calibri');

    // Set Title
    $activeSheet->setCellValue('A1', 'PT. SEKAWAN GLOBAL ENGINEERING');
    $activeSheet->setCellValue('A2', "REPORT DETAIL TS");
    $activeSheet->setCellValue('A3', "PERIODE : " . $period);
    $activeSheet->setCellValue('A4', "Dept: $d");
    $activeSheet->getStyle("A1:A4")->getFont()->setSize(11);

    $activeSheet->fromArray($H2, NULL, "A6");
    $activeSheet->getStyle("A6:J6")->applyFromArray($headerStyle);

    $row_start = 7;
    $row = $row_start;

    $data_loop = array_filter($dataDetail, function ($item) use ($d) {
        return $item['deptid'] == $d;
    });

    $ts_spk = 0;
    $non_spk = 0;
    $unabs = 0;

    foreach ($data_loop as $key => $value) {
        $activeSheet->setCellValue('A' . $row, $value['hourno']);
        $activeSheet->setCellValue('B' . $row, $value['tgl']);
        $activeSheet->setCellValue('C' . $row, $value['nik']);
        $activeSheet->setCellValue('D' . $row, $value['nama']);
        $activeSheet->setCellValue('E' . $row, $value['deptid']);
        $activeSheet->setCellValue('F' . $row, $value['wono']);
        $activeSheet->setCellValue('G' . $row, $value['ts_spk']);
        $activeSheet->setCellValue('H' . $row, $value['non_spk']);
        $activeSheet->setCellValue('I' . $row, $value['ot20'] + $value['rm20']);
        $activeSheet->setCellValue('J' . $row, $value['activity']);

        $activeSheet->getStyle("A$row:J$row")->applyFromArray($innerStyle);
        $row++;
        $ts_spk += $value['ts_spk'];
        $non_spk += $value['non_spk'];
        $unabs += ($value['ot20'] + $value['rm20']);
    }
    $rowTotal = $row;

    $activeSheet->setCellValue('F' . $rowTotal, 'TOTAL');
    $activeSheet->setCellValue('G' . $rowTotal, $ts_spk);
    $activeSheet->setCellValue('H' . $rowTotal, $non_spk);
    $activeSheet->setCellValue('I' . $rowTotal, $unabs);
    $activeSheet->getStyle("F$rowTotal:I$rowTotal")->applyFromArray($footerStyle);

    $activeSheet->getColumnDimension('B')->setWidth(10);
    $activeSheet->getColumnDimension('B')->setWidth(11);
    $activeSheet->getColumnDimension('C')->setWidth(11);
    $activeSheet->getColumnDimension('D')->setWidth(25);
    $activeSheet->getColumnDimension('E')->setWidth(6);
    $activeSheet->getColumnDimension('F')->setWidth(8);
    $activeSheet->getColumnDimension('G')->setWidth(9);
    $activeSheet->getColumnDimension('H')->setWidth(9);
    $activeSheet->getColumnDimension('I')->setWidth(9);
    $activeSheet->getColumnDimension('J')->setWidth(120);
    $n++;

    // Freeze Pane
    $activeSheet->freezePane('A7');
}



// Output
$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
$writer->save($title_document);

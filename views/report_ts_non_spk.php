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
    'NO',
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
$activeSheet->getStyle("A6:H6")->applyFromArray($headerStyle);
$activeSheet->getStyle("J6:N6")->applyFromArray($headerStyle);

// set alignment
$activeSheet->getStyle("A6:F6")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
$activeSheet->getStyle("A6:F6")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$activeSheet->mergeCells("D5:H5");
$activeSheet->setCellValue('D5', 'ALLOCATION OF MAN-HOUR');
$activeSheet->getStyle("D5:H5")->applyFromArray(
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


$activeSheet->mergeCells("J5:N5");
$activeSheet->setCellValue('J5', 'MAN-HOUR CAPACITY');
$activeSheet->getStyle("J5:N5")->applyFromArray(
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
$no = 0;

for ($i = 0; $i < count($dataSummary); $i++) {
    $row = $i + $add;
    $no++;
    $activeSheet->setCellValue('A' . $row, $no);
    $activeSheet->setCellValue('B' . $row, $dataSummary[$i]['nik']);
    $activeSheet->setCellValue('C' . $row, $dataSummary[$i]['nama']);
    $activeSheet->setCellValue('D' . $row, $dataSummary[$i]['deptid']);
    $activeSheet->setCellValue('E' . $row, $dataSummary[$i]['ts_spk']);
    $activeSheet->setCellValue('F' . $row, $dataSummary[$i]['non_spk']);
    $activeSheet->setCellValue('G' . $row, $dataSummary[$i]['unabs']);
    $activeSheet->setCellValue('H' . $row, "=SUM(E$row:G$row)");
    $activeSheet->setCellValue('J' . $row, $dataSummary[$i]['days']);
    $activeSheet->setCellValue('K' . $row, $dataSummary[$i]['hours']);
    $activeSheet->setCellValue('L' . $row, "=(J$row*K$row)");
    $activeSheet->setCellValue('M' . $row, $dataSummary[$i]['ot']);
    $activeSheet->setCellValue('N' . $row, "=SUM(L$row:M$row)");

    $activeSheet->getRowDimension($row)->setOutlineLevel(1);
    $activeSheet->getStyle("A$row:H$row")->applyFromArray($innerStyle);
    $activeSheet->getStyle("J$row:N$row")->applyFromArray($innerStyle);

    if ($i == 0 && isset($dataSummary[$i + 1])) {
        if ($dataSummary[$i]['deptid'] != $dataSummary[$i + 1]['deptid']) {
            $end = $row;
            $totsub = $row + 1;
            $add += 1;
            $activeSheet->setCellValue("D$totsub", $dataSummary[$i]['deptid'] . " Total");
            $activeSheet->setCellValue("E$totsub", "=SUBTOTAL(9,E$start:E$end)");
            $activeSheet->setCellValue("F$totsub", "=SUBTOTAL(9,F$start:F$end)");
            $activeSheet->setCellValue("G$totsub", "=SUBTOTAL(9,G$start:G$end)");
            $activeSheet->setCellValue("H$totsub", "=SUBTOTAL(9,H$start:H$end)");
            $activeSheet->getStyle("A$totsub:H$totsub")->applyFromArray($footerStyle);

            $activeSheet->setCellValue("J$totsub", $dataSummary[$i]['deptid'] . " Total");
            $activeSheet->setCellValue("L$totsub", "=SUBTOTAL(9,L$start:L$end)");
            $activeSheet->setCellValue("M$totsub", "=SUBTOTAL(9,M$start:M$end)");
            $activeSheet->setCellValue("N$totsub", "=SUBTOTAL(9,N$start:N$end)");
            $activeSheet->getStyle("J$totsub:N$totsub")->applyFromArray($footerStyle);


            $start = $i + $add + 1;
            array_push($data_dept, $dataSummary[$i]['deptid']);
        }
    }

    if ($i != 0 && isset($dataSummary[$i + 1])) {
        if ($dataSummary[$i]['deptid'] != $dataSummary[$i + 1]['deptid']) {
            $end = $row;
            $totsub = $row + 1;
            $add += 1;
            $activeSheet->setCellValue("D$totsub", $dataSummary[$i]['deptid'] . " Total");
            $activeSheet->setCellValue("E$totsub", "=SUBTOTAL(9,E$start:E$end)");
            $activeSheet->setCellValue("F$totsub", "=SUBTOTAL(9,F$start:F$end)");
            $activeSheet->setCellValue("G$totsub", "=SUBTOTAL(9,G$start:G$end)");
            $activeSheet->setCellValue("H$totsub", "=SUBTOTAL(9,H$start:H$end)");
            $activeSheet->getStyle("A$totsub:H$totsub")->applyFromArray($footerStyle);

            $activeSheet->setCellValue("J$totsub", $dataSummary[$i]['deptid'] . " Total");
            $activeSheet->setCellValue("L$totsub", "=SUBTOTAL(9,L$start:L$end)");
            $activeSheet->setCellValue("M$totsub", "=SUBTOTAL(9,M$start:M$end)");
            $activeSheet->setCellValue("N$totsub", "=SUBTOTAL(9,N$start:N$end)");
            $activeSheet->getStyle("J$totsub:N$totsub")->applyFromArray($footerStyle);

            $start = $i + $add + 1;
            array_push($data_dept, $dataSummary[$i]['deptid']);
        }
    }

    if ($i == (count($dataSummary) - 1)) {
        $end = $row;
        $totsub = $row + 1;
        $add += 1;
        array_push($data_dept, $dataSummary[$i]['deptid']);
        $activeSheet->setCellValue("D$totsub", $dataSummary[$i]['deptid'] . " Total");
        $activeSheet->setCellValue("E$totsub", "=SUBTOTAL(9,E$start:E$end)");
        $activeSheet->setCellValue("F$totsub", "=SUBTOTAL(9,F$start:F$end)");
        $activeSheet->setCellValue("G$totsub", "=SUBTOTAL(9,G$start:G$end)");
        $activeSheet->setCellValue("H$totsub", "=SUBTOTAL(9,H$start:H$end)");
        $activeSheet->getStyle("A$totsub:H$totsub")->applyFromArray($footerStyle);

        $activeSheet->setCellValue("J$totsub", $dataSummary[$i]['deptid'] . " Total");
        $activeSheet->setCellValue("L$totsub", "=SUBTOTAL(9,L$start:L$end)");
        $activeSheet->setCellValue("M$totsub", "=SUBTOTAL(9,M$start:M$end)");
        $activeSheet->setCellValue("N$totsub", "=SUBTOTAL(9,N$start:N$end)");
        $activeSheet->getStyle("J$totsub:N$totsub")->applyFromArray($footerStyle);
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

$activeSheet->setCellValue('D' . $rowGTotal, 'GRAND TOTAL');
$activeSheet->setCellValue('E' . $rowGTotal, $ts_spk);
$activeSheet->setCellValue('F' . $rowGTotal, $non_spk);
$activeSheet->setCellValue('G' . $rowGTotal, $unabs);
$activeSheet->setCellValue('H' . $rowGTotal, $total);
$activeSheet->getStyle("A$rowGTotal:H$rowGTotal")->applyFromArray($footerStyle);

$activeSheet->setCellValue('J' . $rowGTotal, 'GRAND TOTAL');
$activeSheet->setCellValue('L' . $rowGTotal, $totReg);
$activeSheet->setCellValue('M' . $rowGTotal, $totOT);
$activeSheet->setCellValue('N' . $rowGTotal, $totMHCap);
$activeSheet->getStyle("J$rowGTotal:N$rowGTotal")->applyFromArray($footerStyle);

$rowTitleAbsenTS = $row + 4;
$activeSheet->setCellValue('A' . $rowTitleAbsenTS, "List Karyawan Belum Input TS dan TS Belum di Approve");
$activeSheet->getStyle("A" . $rowTitleAbsenTS)->getFont()->setSize(14);
$activeSheet->getStyle("A" . $rowTitleAbsenTS)->getFont()->setBold(true);

$rowHeaderAbsen = $row + 5;
$row =  $rowHeaderAbsen + 1;
$no = 0;
foreach ($absTS as $key => $value) {

    if ($value['TS'] == 0) {
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

    // if ($value['TS'] != 0) {
    //     $activeSheet->getStyle("F$row")->applyFromArray($footerStyle);
    // }
}
$total_abs = $no * 8;

$activeSheet->setCellValue('J' . $row, 'TOTAL');
$activeSheet->setCellValue('L' . $row, ($total_abs));
$activeSheet->setCellValue('N' . $row, ($total_abs));
$activeSheet->getStyle('J' . $row . ':N' . $row)->applyFromArray($headerStyle);


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

$total_holiday = count($absHoliday) * 8;

$activeSheet->setCellValue('J' . $row, 'TOTAL');
$activeSheet->setCellValue('L' . $row, $total_holiday);
$activeSheet->setCellValue('N' . $row, $total_holiday);
$activeSheet->getStyle('J' . $row . ':N' . $row)->applyFromArray($headerStyle);

$grow = $row + 1;
$activeSheet->setCellValue('J' . $grow , ' GRAND TOTAL');
$activeSheet->setCellValue('L' . $grow , ($totReg + $total_abs + $total_holiday));
$activeSheet->setCellValue('M' . $grow , ($totOT));
$activeSheet->setCellValue('N' . $grow , ($totMHCap + $total_abs + $total_holiday));
$activeSheet->getStyle('J' . $grow  . ':N' . $grow )->applyFromArray($headerStyle);

$activeSheet->getStyle('A' . $rowHeaderAbsenHoliday . ':F' . $rowHeaderAbsenHoliday)->applyFromArray($headerStyle);
$activeSheet->fromArray($arrHeaderAbsenHoliday, NULL, 'A' . $rowHeaderAbsenHoliday);


$rowTitleAbsenTS = $row + 4;
$activeSheet->setCellValue('A' . $rowTitleAbsenTS, "List Karyawan TS Belum Lengkap");
$activeSheet->getStyle("A" . $rowTitleAbsenTS)->getFont()->setSize(14);
$activeSheet->getStyle("A" . $rowTitleAbsenTS)->getFont()->setBold(true);

$rowHeaderTSError = $row + 5;
$row =  $rowHeaderTSError + 1;
$no = 0;
foreach ($absTS as $key => $value) {

    if ($value['TS'] != 0) {
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
}

$activeSheet->getStyle('A' . $rowHeaderTSError . ':F' . $rowHeaderTSError)->applyFromArray($headerStyle);
$activeSheet->fromArray($arrHeaderAbsenTS, NULL, 'A' . $rowHeaderTSError);



// define row size
$row_size = [
    'A' => 11,
    'B' => 14,
    'C' => 25,
    'D' => 14,
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

    $activeSheet->fromArray($H2, NULL, "A7");
    $activeSheet->getStyle("A7:J7")->applyFromArray($headerStyle);
    
    $activeSheet->setCellValue('G2', "TS SPK");
    $activeSheet->setCellValue('H2', "NON SPK");
    $activeSheet->setCellValue('F3', "SALES");
    $activeSheet->setCellValue('F4', "WRTY");
    $activeSheet->setCellValue('F5', "INQ");
    $activeSheet->getStyle("F2:H2")->applyFromArray($headerStyle);

    $row_start = 8;
    $row = $row_start;

    $data_loop = array_filter($dataDetail, function ($item) use ($d) {
        return $item['deptid'] == $d;
    });

    $ts_spk = 0;
    $non_spk = 0;
    $unabs = 0;

    $ts_sls = 0;
    $ts_sls_nspk = 0;

    $ts_wrty = 0;
    $ts_wrty_nspk = 0;
    
    $ts_inq = 0;
    $ts_inq_nspk = 0;

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


        if (substr($value['wono'], 2, 1) == 1 || substr($value['wono'], 2, 1) == 2) {
            $ts_sls += $value['ts_spk'];
            $ts_sls_nspk += $value['non_spk'];
        } else if (substr($value['wono'], 2, 1) == 6) {
            $ts_inq += $value['ts_spk'];
            $ts_inq_nspk +=  $value['non_spk'];
        } else if (substr($value['wono'], 2, 1) == 3 || substr($value['wono'], 2, 1) == 4) {
            $ts_wrty += $value['ts_spk'];
            $ts_wrty_nspk += $value['non_spk'];
        }
    }
    $rowTotal = $row;

    $activeSheet->setCellValue('G3', $ts_sls);
    $activeSheet->setCellValue('G4', $ts_wrty);
    $activeSheet->setCellValue('G5', $ts_inq);

    $activeSheet->setCellValue('H3', $ts_sls_nspk);
    $activeSheet->setCellValue('H4', $ts_wrty_nspk);
    $activeSheet->setCellValue('H5', $ts_inq_nspk);

    $activeSheet->getStyle("F3:H5")->applyFromArray($innerStyle);

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
    $activeSheet->getColumnDimension('G')->setWidth(10);
    $activeSheet->getColumnDimension('H')->setWidth(9);
    $activeSheet->getColumnDimension('I')->setWidth(9);
    $activeSheet->getColumnDimension('J')->setWidth(120);
    $n++;

    // Freeze Pane
    $activeSheet->freezePane('A8');
}

$spreadsheet->setActiveSheetIndex(0);

// Output
$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
$writer->save($title_document);

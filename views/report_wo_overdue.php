<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

function setDimension($activeSheet, $cell, $width)
{
    $activeSheet->getColumnDimension($cell)->setWidth($width);
}

$arrHeaderSummary = [
    'No',
    'WO',
    'Name',
    'Customer',
    'Start Date',
    'Est. Finish Date',
    'Days Left',
    'Category',
    'Type',
    'Marketing',
    'Status',
    'PIC',
];

// Create new Spreadsheet object
$spreadsheet = new Spreadsheet();
$activeSheet = $spreadsheet->getActiveSheet();
$activeSheet->setTitle('Labour');
$spreadsheet->getDefaultStyle()->getFont()->setName('Calibri');
$activeSheet->setCellValue('A1', "PT. SEKAWAN GLOBAL ENGINEERING");
$activeSheet->setCellValue('A2', "Report: WO Over Due ");
$activeSheet->setCellValue('A3', "Periode: As of " . date('d F Y'));
$activeSheet->getStyle("A1:A3")->getFont()->setSize(14);
$activeSheet->getStyle("A1:A3")->getFont()->setBold(true);

//output headers
$headerStyle = [
    'fill' => [
        'fillType' => Fill::FILL_GRADIENT_LINEAR,
        'rotation' => 90,
        'startColor' => [
            'argb' => 'EEEEEEEE',
        ],
        'endColor' => [
            'argb' => 'EEEEEEEE',
        ],
    ],
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => '00000000'],
        ],
    ],
    'font'  => [
        'bold'  =>  true,
        'size' => 11
    ]
];

$activeSheet->getStyle('A5:L5')->applyFromArray($headerStyle);
$activeSheet->fromArray($arrHeaderSummary, NULL, 'A5');
$activeSheet->freezePane('A6');

/* set width */
setDimension($activeSheet, 'A', 5);
setDimension($activeSheet, 'C', 50);
setDimension($activeSheet, 'D', 40);
setDimension($activeSheet, 'E', 17);
setDimension($activeSheet, 'F', 17);
setDimension($activeSheet, 'H', 13);
setDimension($activeSheet, 'J', 18);
setDimension($activeSheet, 'K', 10);
setDimension($activeSheet, 'L', 18);


// Looping data
$row = 0;
$no = 0;
foreach ($data as $key => $value) {
    ++$no;
    $row = (int) $key + 6;

    switch ($value['jenis']) {
        case 1:
            $category = 'Sales';
            break;
        case 2:
            $category = 'Non Warranty';
            break;
        case 3:
            $category = 'Warranty';
            break;
        case 4:
            $category = 'Inquiry';
            break;
        case 5:
            $category = 'Internal';
            break;
        default:
            $category = '';
            break;
    }

    if ($value['approve_extend'] == '') {
        $number = 1;
        $num = $value['conter'];
        $sel = $num - $number;
        if ($sel == 0) {
            $count = 'Rescheduling';
        } else {
            $count = 'Rescheduling (' . $sel . ')';
        }
        $status = $count;
    } elseif ($value['approve_close'] == 'Y') {
        $status = 'Closed';
    } elseif ($value['approve_extend'] == 'Y') {
        $ss = 'Open (' . $value['conter'] . ')';
        $status = $ss;
    } elseif ($value['selesai'] < date('Y-m-d') && $value['approve_close'] == 'N') {
        $status = 'Over Due';
    } elseif ($value['approve_close'] == '') {
        $status = 'Closing';
    } else {
        $status = 'Open';
    }

    if ($value['selesai'] == '1899-12-30') {
        $selesai = '';
    } else {
        $selesai = $value['selesai'];
    }

    if ($value['left'] > 0) {
        $color = "FFFF66";
    } else {
        $color = "FF3333";
    }
    // var_dump($color);die;

    $activeSheet->setCellValue('A' . $row, $no);
    $activeSheet->setCellValue('B' . $row, $value['wono']);
    $activeSheet->setCellValue('C' . $row, $value['desc']);
    $activeSheet->setCellValue('D' . $row, $value['nama']);
    $activeSheet->setCellValue('E' . $row, date('d F Y', strtotime($value['date'])));
    $activeSheet->setCellValue('F' . $row, date('d F Y', strtotime($selesai)));
    $activeSheet->setCellValue('G' . $row, $value['left']);
    $activeSheet->setCellValue('H' . $row, $category);
    $activeSheet->setCellValue('I' . $row, $value['type']);
    $activeSheet->setCellValue('J' . $row, $value['mkt']);
    $activeSheet->setCellValue('K' . $row, $status);
    $activeSheet->setCellValue('L' . $row, $value['pl']);
    /* Style */
    $activeSheet->getStyle("A$row:L$row")->applyFromArray(
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

    $activeSheet->getStyle("G$row")->applyFromArray(
        [
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startColor' => [
                    'argb' => $color,
                ],
                'endColor' => [
                    'argb' => $color,
                ]
            ]
        ]
    );
}

// print_r($spreadsheet);
// die;

$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
$writer->save($title_document);

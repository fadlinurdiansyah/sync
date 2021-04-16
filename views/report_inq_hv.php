<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

function setDimension($activeSheet, $cell, $width)
{
    $activeSheet->getColumnDimension($cell)->setWidth($width);
}

$arrHeaderSummary = [
    'No.',
    'WONO',
    'Project Name',
    'Customer',
    'Start Date',
    'Est. Finish Date',
    'Weeks',
    'Total by Hours',
    'Total by Amount',
    'P. Manager',
    'Marketing'
];

// Create new Spreadsheet object
$spreadsheet = new Spreadsheet();
$activeSheet = $spreadsheet->getActiveSheet();
$activeSheet->setTitle('Inquiry');
$spreadsheet->getDefaultStyle()->getFont()->setName('Calibri');
$activeSheet->setCellValue('A1', "PT. SEKAWAN GLOBAL ENGINEERING");
$activeSheet->setCellValue('A2', "Report: WO Inquiry More Than 72 Hours");
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

$activeSheet->getStyle('A5:K5')->applyFromArray($headerStyle);
$activeSheet->fromArray($arrHeaderSummary, NULL, 'A5');
$activeSheet->freezePane('A6');


// Looping data

$row = 0;
$no = 1;
foreach ($data as $key => $value) {
    $row = (int) $key + 6;
    $activeSheet->setCellValue('A' . $row, $no);
    $activeSheet->setCellValue('B' . $row, $value['wono']);
    $activeSheet->setCellValue('C' . $row, $value['desc']);
    $activeSheet->setCellValue('D' . $row, $value['customer']);
    $activeSheet->setCellValue('E' . $row, $value['date']);
    $activeSheet->setCellValue('F' . $row, $value['selesai']);
    $activeSheet->setCellValue('G' . $row, '=ROUND(DATEDIF(E' . $row . ',F' . $row . ',"D")/7,0)');
    $activeSheet->setCellValue('H' . $row, $value['ts_hour']);
    $activeSheet->setCellValue('I' . $row, $value['ts_rp']);
    $activeSheet->setCellValue('J' . $row, $value['mng_nama']);
    $activeSheet->setCellValue('K' . $row, $value['mkt']);
    $activeSheet->getStyle("A$row:K$row")->applyFromArray(
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

    $activeSheet->getStyle("H$row:I$row")
        ->getNumberFormat()
        ->setFormatCode(NumberFormat::FORMAT_ACCOUNTING_IDR_WITH_COMMA);
    $no++;
}


/* set width */
setDimension($activeSheet, 'A', 4);
setDimension($activeSheet, 'B', 8);
setDimension($activeSheet, 'C', 47);
setDimension($activeSheet, 'D', 47);
setDimension($activeSheet, 'E', 11);
setDimension($activeSheet, 'F', 15);
setDimension($activeSheet, 'G', 7);
setDimension($activeSheet, 'H', 14);
setDimension($activeSheet, 'I', 17);
setDimension($activeSheet, 'J', 17);
setDimension($activeSheet, 'K', 17);


$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
$writer->save($title_document);

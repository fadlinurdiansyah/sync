<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

$arrHeaderSummary = [
    'Item Code',
    'Name',
    'Specification',
    'maker',
    'UOM',
    'Buffer Stock',
    'Reorder Point',
    'Balance Stock',
    'PR Outstanding',
    'PO Outstanding',
    'Order Advice',
];

// Create new Spreadsheet object
$spreadsheet = new Spreadsheet();
$activeSheet = $spreadsheet->getActiveSheet();
$activeSheet->setTitle('Labour');
$spreadsheet->getDefaultStyle()->getFont()->setName('Calibri');
$activeSheet->setCellValue('A1', "PT. SEKAWAN GLOBAL ENGINEERING");
$activeSheet->setCellValue('A2', "Report: SIC (Statistical Inventory Control) Stock");
$activeSheet->setCellValue('A3', "Periode: As of ".date('d F Y'));
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
foreach ($data as $key => $value) {
    $row = (int) $key + 6;
    $activeSheet->setCellValue('A' . $row, $value['stcd']);
    $activeSheet->setCellValue('B' . $row, $value['nama']);
    $activeSheet->setCellValue('C' . $row, $value['spek']);
    $activeSheet->setCellValue('D' . $row, $value['maker']);
    $activeSheet->setCellValue('E' . $row, $value['uom']);
    $activeSheet->setCellValue('F' . $row, $value['buffer']);
    $activeSheet->setCellValue('G' . $row, $value['reorder']);
    $activeSheet->setCellValue('H' . $row, $value['bal_stock']);
    $activeSheet->setCellValue('I' . $row, $value['prouts']);
    $activeSheet->setCellValue('J' . $row, $value['poouts']);
    $activeSheet->setCellValue('K' . $row, $value['advice']);
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
}

$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
$writer->save($title_document);

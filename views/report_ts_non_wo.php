<?php

// var_dump($absTS);
// die;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

function setDimension($activeSheet, $cell, $width)
{
    $activeSheet->getColumnDimension($cell)->setWidth($width);
}

$arrHeaderSummary = [
    'No.',
    'No. TRX',
    'Date',
    'NIK',
    'Name',
    'Dept ID',
    'WONO',
    'Activity',
    'Duration',
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

// Create new Spreadsheet object
$spreadsheet = new Spreadsheet();
$activeSheet = $spreadsheet->getActiveSheet();
$activeSheet->setTitle('TS NON WO');
$spreadsheet->getDefaultStyle()->getFont()->setName('Calibri');
$activeSheet->setCellValue('A1', "PT. SEKAWAN GLOBAL ENGINEERING");
$activeSheet->setCellValue('A2', "Report: TS NON WO ");
$activeSheet->setCellValue('A3', "Periode: As of " . date('d F Y', strtotime('yesterday')));
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

$activeSheet->getStyle('A5:I5')->applyFromArray($headerStyle);
$activeSheet->fromArray($arrHeaderSummary, NULL, 'A5');
$activeSheet->freezePane('A6');

/* set width */
setDimension($activeSheet, 'A', 3);
setDimension($activeSheet, 'B', 10);
setDimension($activeSheet, 'C', 26);
setDimension($activeSheet, 'D', 11);
setDimension($activeSheet, 'E', 25);
setDimension($activeSheet, 'F', 8);
setDimension($activeSheet, 'G', 8);
setDimension($activeSheet, 'H', 100);
setDimension($activeSheet, 'I', 8);


// Looping data

$row = 0;
$no = 0;
foreach ($data as $key => $value) {
    $row = (int) $key + 6;
    $no++;
    $activeSheet->setCellValue('A' . $row, $no);
    $activeSheet->setCellValue('B' . $row, $value['hourno']);
    $activeSheet->setCellValue('D' . $row, $value['nik']);
    $activeSheet->setCellValue('E' . $row, $value['nama']);
    $activeSheet->setCellValue('F' . $row, $value['departemen']);
    $activeSheet->setCellValue('C' . $row, $value['tgl']);
    $activeSheet->setCellValue('G' . $row, $value['wono']);
    $activeSheet->setCellValue('H' . $row, $value['activity']);
    $activeSheet->setCellValue('I' . $row, $value['duration']);
    $activeSheet->getStyle("A$row:I$row")->applyFromArray(
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

$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
$writer->save($title_document);

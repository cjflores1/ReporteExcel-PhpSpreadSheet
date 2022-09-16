<?php

include './headerSpreadSheet.php';

// Set cell A1 with a string value
// $sheet->setCellValue('A1', 'PhpSpreadsheet');

// // Set cell A2 with a numeric value
// $sheet->setCellValue('A2', 12345.6789);

// // Set cell A3 with a boolean value
// $sheet->setCellValue('A3', TRUE);

// // Set cell A4 with a formula
// $sheet->setCellValue(
//     'A4',
//     '=IF(A3, CONCATENATE(A1, " ", A2), CONCATENATE(A2, " ", A1))'
// );

// $sheet->getCell('B8')
//     ->setValue('Some value');

// $writer = new Xlsx($spreadsheet);
// $writer->save('hello world.xlsx');

$inputFileName = './report.xls';
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();

//$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
$spreadsheet = $reader->load($inputFileName);


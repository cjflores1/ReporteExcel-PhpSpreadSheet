<?php

require './vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

$spreadsheet = new Spreadsheet();
/*****************************************************/
/*****************************************************/

// $spreadsheet->createSheet();
// // Create a new worksheet called "My Data"
// $myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, 'Acta');

// Attach the "My Data" worksheet as the first worksheet in the Spreadsheet object
//$spreadsheet->addSheet($myWorkSheet, 0);

$spreadsheet->getProperties()->setCreator("UPRI");
$spreadsheet->getProperties()->setLastModifiedBy("UPRI");
$spreadsheet->getProperties()->setTitle("Acta de Calificaciones");
$spreadsheet->getProperties()->setSubject("Reporte de Acta de Calificaciones");
//$spreadsheet->getProperties()->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.");
//$spreadsheet->getProperties()->setKeywords("office 2007 openxml php");
//$spreadsheet->getProperties()->setCategory("Test result file");


$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("Acta");
//$sheet =$spreadsheet->addSheet($myWorkSheet, 2);

$sheet->getProtection()->setSheet(true);

$styleArray = [
    'font' => [
        'bold' => true,
        'size' => 16,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
];

$styleArray1 = [
    'font' => [
        'bold' => true,
        'size' => 12,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
];

$styleArray2 = [
    'font' => [
        'bold' => true,
        'size' => 10,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
];

$textBold = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
    ],
];

$textBold1 = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
    ],
    'borders' => [
        'outline' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            'color' => ['argb' => 'FF000000'],
        ],
    ],
];

$text1 = [
    'font' => [
        'bold' => false,
        'size' => 10,
    ],
    'alignment' => [
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
    ],
    'borders' => [
        'outline' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            'color' => ['argb' => 'FF000000'],
        ],
    ],
];

$text2 = [
    'font' => [
        'bold' => false,
        'size' => 10,
    ],
    'alignment' => [
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
    'borders' => [
        'outline' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            'color' => ['argb' => 'FF000000'],
        ],
    ],
];

$spreadsheet->getDefaultStyle()->getFont()->setName('Times New Roman');



$sheet->getColumnDimension('A')->setWidth(5);
$sheet->getColumnDimension('B')->setWidth(20);
$sheet->getColumnDimension('C')->setWidth(15);
$sheet->getColumnDimension('D')->setWidth(15);
$sheet->getColumnDimension('E')->setWidth(12);
$sheet->getColumnDimension('F')->setWidth(15);
$sheet->getColumnDimension('G')->setWidth(22);
$sheet->getColumnDimension('H')->setWidth(15);


/*****************************************************/
/*****************************************************/

$nombreArchivo = "./archivosExcel/" . $_GET["name"];

$documento = IOFactory::load($nombreArchivo);

$totalHojas = $documento->getSheetCount();

//para varias hojas
// for($indice = 0; $indice < $totalHojas; $indice++){
//     $hojaActual = $documento->getSheet($indice);
// }

$hojaActual = $documento->getSheet(0);
//echo "<table border=1><tr><th>No.</th><th>APELLIDOS Y NOMBRES</th><th>CI</th><th>RU</th></tr>";
//total de filas
$numeroFilas = $hojaActual->getHighestDataRow();
//total de columnas
$letra = $hojaActual->getHighestColumn();
$contador = 1;
$contadorFilas = 0;
$saltoHoja = -1;
$saltoHojaCuerpo = 0;

for ($indiceFila = 1; $indiceFila <= $numeroFilas; $indiceFila++) {
    $valor = $hojaActual->getCellByColumnAndRow(3, $indiceFila);
    $ci = $hojaActual->getCellByColumnAndRow(33, $indiceFila);
    $ru = $hojaActual->getCellByColumnAndRow(47, $indiceFila);
    
    if (strcmp($valor, "") != 0) {
        
        $nota = 0;

        if($contador==1 || ($contador-1)%33==0){
            //echo "<tr><td colspan=4></td></tr>";
            $sheet->setCellValue('A'.($contador+$saltoHoja+1), 'UNIVERSIDAD MAYOR DE SAN ANDRES');
            $sheet->mergeCells('A'.($contador+$saltoHoja+1).':H'.($contador+$saltoHoja+1));
            $sheet->getStyle('A'.($contador+$saltoHoja+1))->applyFromArray($styleArray);
            $sheet->setCellValue('A'.($contador+$saltoHoja+2), 'FACULTAD DE DERECHO Y CIENCIAS POLITICAS');
            $sheet->mergeCells('A'.($contador+$saltoHoja+2).':H'.($contador+$saltoHoja+2));
            $sheet->getStyle('A'.($contador+$saltoHoja+2))->applyFromArray($styleArray1);
            $sheet->setCellValue('A'.($contador+$saltoHoja+3), 'UNIDAD DE POSTGRADO Y RELACIONES INTERNACIONALES');
            $sheet->mergeCells('A'.($contador+$saltoHoja+3).':H'.($contador+$saltoHoja+3));
            $sheet->getStyle('A'.($contador+$saltoHoja+3))->applyFromArray($styleArray2);
            $sheet->setCellValue('A'.($contador+$saltoHoja+5), 'ACTA DE CALIFICACIONES');
            $sheet->mergeCells('A'.($contador+$saltoHoja+5).':H'.($contador+$saltoHoja+5));
            $sheet->getStyle('A'.($contador+$saltoHoja+5))->applyFromArray($styleArray);
    
            $sheet->setCellValue('B'.($contador+$saltoHoja+7), 'AREA: ')->getStyle('B'.($contador+$saltoHoja+7))->applyFromArray($textBold);
            $sheet->setCellValue('B'.($contador+$saltoHoja+8), 'NIVEL: ')->getStyle('B'.($contador+$saltoHoja+8))->applyFromArray($textBold);
            $sheet->setCellValue('G'.($contador+$saltoHoja+8), 'GESTION: ')->getStyle('G'.($contador+$saltoHoja+8))->applyFromArray($textBold);
            $sheet->getRowDimension(($contador+$saltoHoja+9))->setRowHeight(28);
            $sheet->setCellValue('B'.($contador+$saltoHoja+9), 'PROGRAMA: ')->getStyle('B'.($contador+$saltoHoja+9))->applyFromArray($textBold);
            $sheet->setCellValue('G'.($contador+$saltoHoja+9), 'VERSION: ')->getStyle('G'.($contador+$saltoHoja+9))->applyFromArray($textBold);
            $sheet->getRowDimension(($contador+$saltoHoja+10))->setRowHeight(28);
            $sheet->setCellValue('B'.($contador+$saltoHoja+10), 'MODULO: ')->getStyle('B'.($contador+$saltoHoja+10))->applyFromArray($textBold);
            $sheet->setCellValue('G'.($contador+$saltoHoja+10), 'FECHA: ')->getStyle('G'.($contador+$saltoHoja+10))->applyFromArray($textBold);
            $sheet->setCellValue('B'.($contador+$saltoHoja+11), 'CATEDRATICO: ')->getStyle('B'.($contador+$saltoHoja+11))->applyFromArray($textBold);
    
            $sheet->getRowDimension(($contador+$saltoHoja+13))->setRowHeight(28);
            $sheet->setCellValue('A'.($contador+$saltoHoja+13), 'No.')->getStyle('A'.($contador+$saltoHoja+13))->applyFromArray($textBold1);
            $sheet->setCellValue('B'.($contador+$saltoHoja+13), 'APELLIDOS Y NOMBRES')->getStyle('B'.($contador+$saltoHoja+13).':D'.($contador+$saltoHoja+13))->applyFromArray($textBold1);
            $sheet->mergeCells('B'.($contador+$saltoHoja+13).':D'.($contador+$saltoHoja+13));
            $sheet->setCellValue('E'.($contador+$saltoHoja+13), 'CEDULA DE IDENTIDAD')->getStyle('E'.($contador+$saltoHoja+13))->applyFromArray($textBold1)->getAlignment()->setWrapText(true);
            $sheet->setCellValue('F'.($contador+$saltoHoja+13), 'CALIFICACION NUMERAL')->getStyle('F'.($contador+$saltoHoja+13))->applyFromArray($textBold1)->getAlignment()->setWrapText(true);
            $sheet->getRowDimension(($contador+$saltoHoja+10))->setRowHeight(25);
            $sheet->setCellValue('G'.($contador+$saltoHoja+13), 'CALIFICACION LITERAL')->getStyle('G'.($contador+$saltoHoja+13))->applyFromArray($textBold1)->getAlignment()->setWrapText(true);
            $sheet->getRowDimension(($contador+$saltoHoja+10))->setRowHeight(25);
            $sheet->setCellValue('H'.($contador+$saltoHoja+13), 'PONDERACION')->getStyle('H'.($contador+$saltoHoja+13))->applyFromArray($textBold1);
            $sheet->getStyle('A'.($contador+$saltoHoja+13).':H'.($contador+$saltoHoja+13))->applyFromArray($styleArray2);

            $contadorFilas += 14;
            $saltoHoja+=14;  
        }
        
        //echo "<tr><td>" . $contador . "</td><td>" . $valor. "-" .$saltoHoja. "</td><td>" . $ci . "</td><td>" . $ru . "</td></tr>";

        $sheet->setCellValue('A'.($contadorFilas+$saltoHojaCuerpo), $contador)->getStyle('A'.($contadorFilas+$saltoHojaCuerpo))->applyFromArray($text2);
        $sheet->setCellValue('B'.($contadorFilas+$saltoHojaCuerpo), $valor)->getStyle('B'.($contadorFilas+$saltoHojaCuerpo).':D'.($contadorFilas+$saltoHojaCuerpo))->applyFromArray($text1);
        $sheet->mergeCells('B'.($contadorFilas+$saltoHojaCuerpo).':D'.($contadorFilas+$saltoHojaCuerpo));
        $sheet->setCellValue('E'.($contadorFilas+$saltoHojaCuerpo), $ci)->getStyle('E'.($contadorFilas+$saltoHojaCuerpo))->applyFromArray($text2);
        $sheet->setCellValue('F'.($contadorFilas+$saltoHojaCuerpo), $nota)->getStyle('F'.($contadorFilas+$saltoHojaCuerpo))->applyFromArray($text2)->getAlignment()->setWrapText(true);
        $sheet->getStyle('F'.($contadorFilas+$saltoHojaCuerpo))->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);
        $sheet->getStyle('F'.($contadorFilas+$saltoHojaCuerpo))->getProtection()->setLocked(\PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_UNPROTECTED);

        $literales1 = 'IF(F'.($contadorFilas+$saltoHojaCuerpo).'=0,"CERO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=1,"UNO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=2,"DOS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=3,"TRES",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=4,"CUATRO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=5,"CINCO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=6,"SEIS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=7,"SIETE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=8,"OCHO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=9,"NUEVE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=10,"DIEZ",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=11,"ONCE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=12,"DOCE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=13,"TRECE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=14,"CATORCE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=15,"QUINCE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=16,"DIECISEIS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=17,"DIECISIETE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=18,"DIECIOCHO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=19,"DIECINUEVE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=20,"VEINTE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=21,"VEINTIUNO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=22,"VEINTIDOS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=23,"VEINTITRES",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=24,"VEINTICUATRO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=25,"VEINTICINCO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=26,"VEINTISEIS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=27,"VEINTISIETE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=28,"VEINTIOCHO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=29,"VEINTINUEVE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=30,"TREINTA",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=31,"TREINTA Y UNO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=32,"TREINTA Y DOS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=33,"TREINTA Y TRES",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=34,"TREINTA Y CUATRO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=35,"TREINTA Y CINCO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=36,"TREINTA Y SEIS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=37,"TREINTA Y SIETE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=38,"TREINTA Y OCHO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=39,"TREINTA Y NUEVE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=40,"CUARENTA",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=41,"CUARENTA Y UNO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=42,"CUARENTA Y DOS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=43,"CUARENTA Y TRES",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=44,"CUARENTA Y CUATRO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=45,"CUARENTA Y CINCO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=46,"CUARENTA Y SEIS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=47,"CUARENTA Y SIETE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=48,"CUARENTA Y OCHO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=49,"CUARENTA Y NUEVE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=50,"CINCUENTA",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=51,"CINCUENTA Y UNO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=52,"CINCUENTA Y DOS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=53,"CINCUENTA Y TRES",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=54,"CINCUENTA Y CUATRO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=55,"CINCUENTA Y CINCO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=56,"CINCUENTA Y SEIS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=57,"CINCUENTA Y SIETE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=58,"CINCUENTA Y OCHO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=59,"CINCUENTA Y NUEVE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=60,"SESENTA","")))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))';
        $literales2 = 'IF(F'.($contadorFilas+$saltoHojaCuerpo).'=61,"SESENTA Y UNO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=62,"SESENTA Y DOS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=63,"SESENTA Y TRES",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=64,"SESENTA Y CUATRO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=65,"SESENTA Y CINCO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=66,"SESENTA Y SEIS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=67,"SESENTA Y SIETE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=68,"SESENTA Y OCHO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=69,"SESENTA Y NUEVE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=70,"SETENTA",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=71,"SETENTA Y UNO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=72,"SETENTA Y DOS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=73,"SETENTA Y TRES",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=74,"SETENTA Y CUATRO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=75,"SETENTA Y CINCO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=76,"SETENTA Y SEIS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=77,"SETENTA Y SIETE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=78,"SETENTA Y OCHO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=79,"SETENTA Y NUEVE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=80,"OCHENTA",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=81,"OCHENTA Y UNO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=82,"OCHENTA Y DOS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=83,"OCHENTA Y TRES",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=84,"OCHENTA Y CUATRO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=85,"OCHENTA Y CINCO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=86,"OCHENTA Y SEIS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=87,"OCHENTA Y SIETE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=88,"OCHENTA Y OCHO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=89,"OCHENTA Y NUEVE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=90,"NOVENTA",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=91,"NOVENTA Y UNO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=92,"NOVENTA Y DOS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=93,"NOVENTA Y TRES",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=94,"NOVENTA Y CUATRO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=95,"NOVENTA Y CINCO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=96,"NOVENTA Y SEIS",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=97,"NOVENTA Y SIETE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=98,"NOVENTA Y OCHO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=99,"NOVENTA Y NUEVE",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=100,"CIEN","-"))))))))))))))))))))))))))))))))))))))))';     
        
        $sheet->setCellValue('X'.($contadorFilas+$saltoHojaCuerpo), $ru);
        $sheet->setCellValue('Y'.($contadorFilas+$saltoHojaCuerpo), '=IF(AND(F'.($contadorFilas+$saltoHojaCuerpo).'>=0,F'.($contadorFilas+$saltoHojaCuerpo).'<61),'.$literales1.',"ERROR")');
        $sheet->setCellValue('Z'.($contadorFilas+$saltoHojaCuerpo), '=IF(AND(F'.($contadorFilas+$saltoHojaCuerpo).'>60,F'.($contadorFilas+$saltoHojaCuerpo).'<101),'.$literales2.',"ERROR")');
    
        $sheet->setCellValue('G'.($contadorFilas+$saltoHojaCuerpo), '=IF(F'.($contadorFilas+$saltoHojaCuerpo).'<61,Y'.($contadorFilas+$saltoHojaCuerpo).',Z'.($contadorFilas+$saltoHojaCuerpo).')')->getStyle('G'.($contadorFilas+$saltoHojaCuerpo))->applyFromArray($text2)->getAlignment()->setWrapText(true);
        
        $sheet->getColumnDimension('X')->setVisible(false);
        $sheet->getColumnDimension('Y')->setVisible(false);
        $sheet->getColumnDimension('Z')->setVisible(false);
        
        $ponderacion = '=IF(ISNUMBER(F'.($contadorFilas+$saltoHojaCuerpo).'),IF(AND(F'.($contadorFilas+$saltoHojaCuerpo).'>90, F'.($contadorFilas+$saltoHojaCuerpo).'<101),"EXCELENTE", IF(AND(F'.($contadorFilas+$saltoHojaCuerpo).'>80, F'.($contadorFilas+$saltoHojaCuerpo).'<91),"MUY BUENO",IF(AND(F'.($contadorFilas+$saltoHojaCuerpo).'>70, F'.($contadorFilas+$saltoHojaCuerpo).'<81),"BUENO",IF(AND(F'.($contadorFilas+$saltoHojaCuerpo).'>65, F'.($contadorFilas+$saltoHojaCuerpo).'<71),"APROBADO",IF(AND(F'.($contadorFilas+$saltoHojaCuerpo).'>0, F'.($contadorFilas+$saltoHojaCuerpo).'<66),"REPROBADO",IF(F'.($contadorFilas+$saltoHojaCuerpo).'=0,"N.S.P.","ERROR")))))),"ERROR")';
                
        $sheet->setCellValue('H'.($contadorFilas+$saltoHojaCuerpo), $ponderacion)->getStyle('H'.($contadorFilas+$saltoHojaCuerpo))->applyFromArray($text2);        

        $contador++;
        $contadorFilas++;
    }
}

//$spreadsheet->removeSheetByIndex(1);


//$writer = new Xlsx($spreadsheet);
//$writer->save('ActaCalificaciones.xlsx');

$nombreDelDocumento = "ActaCalificaciones.xlsx";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="' . $nombreDelDocumento . '"');
header('Cache-Control: max-age=0');
 
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');
exit;

<?php
    // require 'vendor/autoload.php';

    // use PhpOffice\PhpSpreadsheet\Spreadsheet;
    // use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
 
    //require 'vendor/autoload.php';
    //require('vendor/setasign/fpdf/fpdf.php');
    require './fpdf/fpdf.php';

    $pdf = new FPDF();
    $pdf->AddPage();
    $pdf->SetFont('Arial','B',16);
    $pdf->Cell(40,10,utf8_decode('¡Hola, Mundo!'));
    $pdf->Output();

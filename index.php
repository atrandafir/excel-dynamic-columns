<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Load demo data
$rows = require 'demo-data.php';

// === Example 1: Normal way of naming Excel columns ===
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Normal way of naming Excel columns');

$i=3;

// Header
$sheet->setCellValue('A'.$i, 'Name');
$sheet->setCellValue('B'.$i, 'Phone');
$sheet->setCellValue('C'.$i, 'City');

// Rows
foreach ($rows as $row) {
  $i++;
  $sheet->setCellValue('A'.$i, $row['name']);
  $sheet->setCellValue('B'.$i, $row['phone']);
  $sheet->setCellValue('C'.$i, $row['city']);
}

// Styling
$lastRowIndex=$i;

$sheet->getStyle('A3:C'.$lastRowIndex)->applyFromArray([
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
]);

$sheet->getStyle('A3:C3')->getFill()
    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
    ->getStartColor()->setARGB('FFA0A0A0');
    
$sheet->getColumnDimension('A')->setWidth(12);
$sheet->getColumnDimension('B')->setWidth(25);
$sheet->getColumnDimension('C')->setWidth(25);

$writer = new Xlsx($spreadsheet);
$writer->save('normal.xlsx');


// === Example 2: Dynamic way of naming Excel columns ===

// Load the column names helper
require 'ExCol.php';

ExCol::reset(); // reset mapping before using

$show_company_column=true;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Dynamic way of naming Excel columns');

$i=3;

// Header
$sheet->setCellValue(ExCol::get('name', $i), 'Name');
$sheet->setCellValue(ExCol::get('phone', $i), 'Phone');
if ($show_company_column) {
  $sheet->setCellValue(ExCol::get('company', $i), 'Company');
}
$sheet->setCellValue(ExCol::get('city', $i), 'City');

// Rows
foreach ($rows as $row) {
  $i++;
  $sheet->setCellValue(ExCol::get('name', $i), $row['name']);
  $sheet->setCellValue(ExCol::get('phone', $i), $row['phone']);
  if ($show_company_column) {
    $sheet->setCellValue(ExCol::get('company', $i), $row['company']);
  }
  $sheet->setCellValue(ExCol::get('city', $i), $row['city']);
}

$lastColLetter=ExCol::getLast();

// Styling
$lastRowIndex=$i;

$sheet->getStyle('A3:'.$lastColLetter.$lastRowIndex)->applyFromArray([
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
]);

$sheet->getStyle("A3:{$lastColLetter}3")->getFill()
    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
    ->getStartColor()->setARGB('FFA0A0A0');
    
$sheet->getColumnDimension(ExCol::get('name'))->setWidth(12);
if ($show_company_column) {
  $sheet->getColumnDimension(ExCol::get('company'))->setWidth(35);
}
$sheet->getColumnDimension(ExCol::get('phone'))->setWidth(25);
$sheet->getColumnDimension(ExCol::get('city'))->setWidth(25);

$writer = new Xlsx($spreadsheet);
$writer->save('dynamic.xlsx');

// End
echo 'Excels were generated! Look in the folder :-)';
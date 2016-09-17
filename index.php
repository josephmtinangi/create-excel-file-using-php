<?php

require './vendor/autoload.php';

// Create a new PHPExcel object
$mySpreadSheet = new PHPExcel();

// Set document properties
$mySpreadSheet->getProperties()
    ->setCreator("Joseph Mtinangi")
    ->setTitle('EAC Countries');
    
$myWorksheet = $mySpreadSheet->getSheet(0);
$myWorksheet->setTitle('East Africa Community Countries');

$myWorksheet->setCellValue('A1', 'SN');
$myWorksheet->setCellValue('B1', 'Country');
$myWorksheet->setCellValue('C1', 'Capital');

$myWorksheet->getStyle('A1')->getFont()->setBold(true);
$myWorksheet->getStyle('B1')->getFont()->setBold(true);
$myWorksheet->getStyle('C1')->getFont()->setBold(true);

$myWorksheet->setCellValue('A2', '1');
$myWorksheet->setCellValue('B2', 'Tanzania');
$myWorksheet->setCellValue('C2', 'Dodoma');

$myWorksheet->setCellValue('A3', '2');
$myWorksheet->setCellValue('B3', 'Kenya');
$myWorksheet->setCellValue('C3', 'Nairobi');

$myWorksheet->setCellValue('A4', '3');
$myWorksheet->setCellValue('B4', 'Uganda');
$myWorksheet->setCellValue('C4', 'Kampala');

$myWorksheet->setCellValue('A5', '4');
$myWorksheet->setCellValue('B5', 'Rwanda');
$myWorksheet->setCellValue('C5', 'Kigali');

$myWorksheet->setCellValue('A6', '5');
$myWorksheet->setCellValue('B6', 'Burudi');
$myWorksheet->setCellValue('C6', 'Bujumbura');

$writer = PHPExcel_IOFactory::createWriter($mySpreadSheet, 'Excel2007');
$writer->save('eac.xlsx');

echo 'Done writing files. ', EOL;
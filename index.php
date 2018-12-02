<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

/**
 * Книга. Можно загрузить из файла или создать самому.
 * Результат (в памяти) можно копировать.
 * Помимо страниц, содержит метаинфрмацию, права доступа
 */
$spreadsheet = new Spreadsheet();

/**
 * Всегда есть активный лист. Есть методы для смены и получения активного листа,
 * а также получения любого листа по имени (byName) или по индексу
 */
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('One');

foreach ([1, 2, 3] as $rowIndex) {
    foreach (['A', 'B', 'C'] as $columnIndex) {
        $index = $columnIndex . $rowIndex;
        $sheet->setCellValue($index, $index);
    }
}

/**
 * Формулы
 */
$sheet->setCellValue(
    'A4',
    '=IF(A3, CONCATENATE(A1, " ", A2), CONCATENATE(A2, " ", A1))'
);

// не клон! см. имплементацию
$sheetB = $sheet->copy();
$sheetB->setTitle('Two');
$spreadsheet->addSheet($sheetB);

$sheetC = $spreadsheet->createSheet();
$sheetC->setTitle('Three');
$sheetC->setCellValue('A1', 'AAAArgh!!');

$writer = new Xlsx($spreadsheet);
$inputFileName = 'files/hello_world.xlsx';
$writer->save($inputFileName);

/**
 * Вариант с загрузкой
 */
$spreadsheet2 = IOFactory::load($inputFileName);

/**
 * Чтобы правильно удалить книгу из памяти, нужно сначала разбить цикличиские ссылки
 */
$spreadsheet->disconnectWorksheets();
unset($spreadsheet);
<?php

include(dirname(__DIR__).DIRECTORY_SEPARATOR.'vendor/autoload.php');

$excel = new \Vtiful\Kernel\Excel(['path' => __DIR__]);
$excel->openFile('jd1.xlsx')->openSheet();
$sheet_list = $excel->sheetList();
var_dump($sheet_list);
if (empty($sheet_list)) {
    return [];
}
$sheet_name = array_shift($sheet_list);
$data = $excel->openSheet($sheet_name, \Vtiful\Kernel\Excel::SKIP_EMPTY_ROW)
    ->getSheetData();
var_dump($data);
<?php

require "../vendor/autoload.php";

$excel = new Tianmiao\Excel\ExcelZipExport('', [
    'zip_path' => __DIR__ . "/output.zip",
    'single_max_row' => 10000,
]);

//$excel = new Tianmiao\Excel\ExcelZipExport();

//$excel->setZipPath(__DIR__ . "/output.zip");

$excel->setHeader((["费率", "笔数", "充值进账", "手续费", "利润"]));

$data = ([
    ["1.0%", 0, 0.00, 0, 0],
    ["2.0%", 0, 0.00, 0, 0],
    ["3.0%", 0, 0.00, 0, 0],
    ["4.0%", 0, 0.00, 0, 0],
    ["5.0%", 0, 0.00, 0, 0],
    ["6.0%", 0, 0.00, 0, 0],
    ["7.0%", 0, 0.00, 0, 0],
    ["8.0%", 0, 0.00, 0, 0],
    ["9.0%", 0, 0.00, 0, 0],
    ["10.0%", 0, 0.00, 0, 0],
]);
for ($i=1;$i<10000;$i++){
    $excel->write($data);
    var_dump(memory_get_usage(true));
}


$excel->save();
var_dump(memory_get_usage(true));

<?php

require "./vendor/autoload.php";

//$excel = new Tianmiao\Excel\ExcelZipExport('', ['zip_path' => 'www/output.zip']);

$excel = new Tianmiao\Excel\ExcelZipExport();

$excel->setZipPath(__DIR__ . "/output.zip");

$excel->setHead((["费率", "笔数", "充值进账", "手续费", "利润"]));

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
    ["11.0%", 0, 0.00, 0, 0],
    ["12.0%", 0, 0.00, 0, 0],
    ["13.0%", 0, 0.00, 0, 0],
    ["14.0%", 0, 0.00, 0, 0],
]);

$excel->write($data);
$excel->save();

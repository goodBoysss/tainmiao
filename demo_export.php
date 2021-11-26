<?php

require "./vendor/autoload.php";

$excel = new Tianmiao\Excel\ExcelExport();

$excel->write(array("费率", "笔数", "充值进账", "手续费", "利润"));
$excel->write(array(
    array("1.0%", 0, 0.00, 0, 0),
    array("合计", 0, 0.00, 0, 0),
));

$excel->save(__DIR__ . "/output.xls");

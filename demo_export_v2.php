<?php

require "./vendor/autoload.php";
ini_set('default_charset', 'UTF-8');
$excel = new Tianmiao\Excel\ExcelExport();

$excel->load("./支付宝批量付款文件上传模板.xls");
//
$excel->setInsertRow(3);
$excel->write(array(
    1, '17610151971', '1', 1, ""
));
//$excel->write(array("","","","",""));
//$excel->write(array("","","","",""));
//$excel->write(array("","","","",""));
//$excel->write(array("","","","",""));

$excel->save(__DIR__ . "/output.xls");

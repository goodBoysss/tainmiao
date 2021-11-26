<?php

require "./vendor/autoload.php";

$reader = new Tianmiao\Excel\ExcelReader('D:\wamp64\www\tianmiao\financial\storage\template\alipay\支付宝批量付款文件上传模板_result.xls');

var_dump($reader->getRowCount());
var_dump($reader->getFullData());
die();


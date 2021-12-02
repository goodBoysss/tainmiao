<?php

require "./vendor/autoload.php";

$excel = new Tianmiao\Excel\CsvExport();

//$excel->load("./支付宝批量付款文件上传模板.xls");

$excel->write(array("支付宝批量付款文件模板（前面两行请勿删除）"));
$excel->write(array("序号（必填）","收款方支付宝账号（必填）",'收款方姓名（必填）','金额（必填，单位：元）',"备注（选填）"));


for ($i=0;$i<10;$i++){
    $excel->write(array(
        array($i, '17610151971', '你好', 1, ""),
    ));
}

$excel->save(__DIR__ . "/output.csv");

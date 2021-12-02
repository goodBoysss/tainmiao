<?php

require "./vendor/autoload.php";

$excel = new Tianmiao\Excel\ExcelExport();

$excel->load("D:/wamp64/www/tianmiao/financial/storage/template/alipay/支付宝批量付款文件上传模板.xls");
//$excel->load("./output.xls");

$excel->write(array("支付宝批量付款文件模板（前面两行请勿删除）"));
$excel->write(array("序号（必填）","收款方支付宝账号（必填）",'收款方姓名（必填）','金额（必填，单位：元）',"备注（选填）"));
for ($i=0;$i<50;$i++){
    $excel->write(array(
        array($i, '17610151971', '张林晓', 1, ""),
    ));
}


//$excel->merge("A1","E1");
//$excel->center("A1:E1");

$excel->save(__DIR__ . "/output_30_v3.xls");

# excel文件操作
基于phpexcel，对调用方式进行简化，目前支持excel导出


## 环境要求

* PHP >= 5.6


## 安装（composer包）
```shell
composer require tianmiao/excel
```



## 示例
```php

require "./vendor/autoload.php";

$excel = new Tianmiao\Excel\ExcelExport();

$excel->write(array("费率", "笔数", "充值进账", "手续费", "利润"));
$excel->write(array(
    array("1.0%", 0, 0.00, 0, 0),
    array("合计", 0, 0.00, 0, 0),
));

$excel->save(__DIR__ . "/output.xls");

```

<?php

require "./vendor/autoload.php";

$reader = new Tianmiao\Excel\CsvReader("D:/wamp64/www/github/tianmiao/output.csv");

var_dump($reader->getAllData());
die();


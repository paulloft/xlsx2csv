<?php
include __DIR__ . '/vendor/autoload.php';

$converter = new \Utils\XlsxToCsv('test2.xlsx');
$converter->convert('example.csv');

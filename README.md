# xlsx to csv converter
Lightweight and simple XLSX to CSV converter. Works faster and consumes less memory than converting via PHPExcell

Usages
------------
```php
$converter = new \Utils\XlsxToCsv('example.xlsx');
$converter->convert('example.csv');
```

Installation
------------

*Requres PHP 7.3 or higher*

Add `paulloft/xlsx2csv` to your composer.json.

```json
"require": {
    "paulloft/xlsx2csv": "*"
}
```

or run in shell

```shell
composer require paulloft/xlsx2csv
```



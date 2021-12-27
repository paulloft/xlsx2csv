<?php
/**
 * @author PaulLoft <info@paulloft.ru>
 */

namespace Utils;

use Exception;
use SimpleXMLElement;
use XMLReader;
use ZipArchive;
use function array_key_exists;
use function count;
use function dirname;
use function strlen;

class XlsxToCsv
{
    public static $delimeter = ';';
    public static $enclosure = '"';
    public static $escape = '\\';

    public static $formatDateTime = 'd.m.Y H:i';
    public static $formatDate = 'd.m.Y';

    /**
     * @var string[]
     */
    protected const NUMS = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'];
    protected const TYPE_DATETIME = 'datetime';
    protected const TYPE_DATE = 'date';
    protected const TYPE_OTHER = 'other';

    protected const ATTR_TYPE = 't';
    protected const ATTR_ROW = 'r';
    protected const ATTR_STRING = 's';
    protected const ATTR_VALUE = 'v';

    /**
     * @var string
     */
    private $xlsxPath;

    /**
     * @var string
     */
    private $csvPath;

    /**
     * @var string
     */
    private $tmpDir;

    /**
     * @var bool
     */
    private $extracted = false;

    /**
     * XlsxToCsv constructor.
     *
     * @param string $xlsxPath path to xlsx file
     */
    public function __construct(string $xlsxPath)
    {
        $this->xlsxPath = $xlsxPath;
        $this->tmpDir = sys_get_temp_dir() . '/xlsx_to_csv';
    }

    /**
     * @param string $saveCsvPath path to save csv file
     * @param int $sheetNumber
     * @return bool
     * @throws XlsxConverterException
     */
    public function convert(string $saveCsvPath, int $sheetNumber = 1): bool
    {
        $this->unpack();

        $csvDir = dirname($saveCsvPath);
        $this->mkdir($csvDir);

        $csvFile = fopen($saveCsvPath, 'wb');
        if ($csvFile === false) {
            throw new XlsxConverterException("Unable to create csv file on path $saveCsvPath");
        }

        $strings = $this->getSharedStrings();
        $dateTypes = $this->getDateTypes();

        $filename = "$this->tmpDir/xl/worksheets/sheet$sheetNumber.xml";
        $reader = $this->getReader($filename);

        $rowCount = '0';
        while ($reader->read()) {
            if ($reader->name === 'row') {
                break;
            }
        }
        ob_start();

        while ($reader->name === 'row') {
            $thisRow = [];

            $node = $this->getNode($reader);
            $result = $this->xmlObjToArray($node);

            $cells = $result['children']['c'];
            $rowNo = static::getAttr($result, static::ATTR_ROW);
            $colAlpha = 'A';

            foreach ($cells as $cell) {
                if (array_key_exists(static::ATTR_VALUE, $cell['children'])) {
                    $cellNo = str_replace(static::NUMS, '', static::getAttr($cell, static::ATTR_ROW));
                    $value = static::getValue($cell);

                    for ($col = $colAlpha; $col !== $cellNo; $col++) {
                        $thisRow[] = ' ';
                        $colAlpha++;
                    }

                    if (static::getAttr($cell, static::ATTR_TYPE) === static::ATTR_STRING) {
                        $thisRow[] = $strings[$value] ?? $value;
                    } else {
                        $type = static::getAttr($cell, static::ATTR_STRING);
                        $formatType = $dateTypes[$type] ?? static::TYPE_OTHER;

                        switch ($formatType) {
                            case static::TYPE_DATETIME:
                                $thisRow[] = static::excellTimestampToDate((int)$value, static::$formatDateTime);
                                break;

                            case static::TYPE_DATE:
                                $thisRow[] = static::excellTimestampToDate((int)$value, static::$formatDate);
                                break;

                            default:
                                $thisRow[] = $value;
                        }
                    }
                } else {
                    $thisRow[] = '';
                }

                $colAlpha++;
            }

            $rowLength = count($thisRow);
            $rowCount++;
            $emptyRow = [];

            while ($rowCount < $rowNo) {
                for ($c = 0; $c < $rowLength; $c++) {
                    $emptyRow[] = '';
                }

                if (!empty($emptyRow)) {
                    $this->writeArrayToCsv($csvFile, $emptyRow);
                }

                $rowCount++;
            }

            $this->writeArrayToCsv($csvFile, $thisRow);

            $reader->next('row');

            $result = null;
        }

        $reader->close();
        ob_end_flush();

        $this->cleanUp($this->tmpDir);
        fclose($csvFile);

        return true;
    }

    /**
     * @param int $timestamp
     * @param string $format
     * @return string
     */
    public static function excellTimestampToDate(int $timestamp, string $format): string
    {
        return date($format, ($timestamp - 25569) * 86400);
    }

    /**
     * @param string $file
     * @return XMLReader
     */
    protected function getReader(string $file): XMLReader
    {
        $reader = new XMLReader();
        $reader->open($file);

        return $reader;
    }

    /**
     * @param XMLReader $reader
     * @return SimpleXMLElement
     * @throws XlsxConverterException
     */
    protected function getNode(XMLReader $reader): SimpleXMLElement
    {
        try {
            return new SimpleXMLElement($reader->readOuterXml());
        } catch (Exception $exception) {
            throw new XlsxConverterException($exception->getMessage());
        }
    }

    /**
     * @return array
     * @throws XlsxConverterException
     */
    protected function getSharedStrings(): array
    {
        $strings = [];
        $filename = "$this->tmpDir/xl/sharedStrings.xml";

        $reader = $this->getReader($filename);

        while ($reader->read()) {
            if ($reader->name === 'si') {
                break;
            }
        }

        ob_start();

        while ($reader->name === 'si') {
            $node = $this->getNode($reader);
            $result = $this->xmlObjToArray($node);

            $strings[] = static::getValue($result, static::ATTR_TYPE);

            $reader->next('si');
            $result = null;
        }

        ob_end_flush();
        $reader->close();

        return $strings;
    }

    /**
     * get value types of cells
     *
     * @return array
     * @throws XlsxConverterException
     */
    protected function getDateTypes(): array
    {
        $types = [];
        $formats = [];
        $filename = "$this->tmpDir/xl/styles.xml";

        $reader = $this->getReader($filename);

        ob_start();

        while ($reader->read()) {
            if ($reader->name === 'numFmt') {
                $id = $reader->getAttribute('numFmtId');
                $format = $reader->getAttribute('formatCode');
                $formats[$id] = $this->getFormatTypes($format);
            } elseif ($reader->name === 'cellXfs') {
                $node = $this->getNode($reader);
                $elements = $this->xmlObjToArray($node)['children']['xf'] ?? [];

                foreach ($elements as $key => $element) {
                    if (static::getAttr($element, 'applynumberformat')) {
                        $formatID = static::getAttr($element, 'numfmtid');
                        $types[$key] = $formats[$formatID] ?? static::TYPE_OTHER;
                    }
                }
                break;
            }
        }

        ob_end_flush();
        $reader->close();

        return $types;
    }

    /**
     * @param array $element
     * @param string $attr
     * @return string|null
     */
    protected static function getAttr(array $element, string $attr): ?string
    {
        return $element['attributes'][$attr] ?? null;
    }

    /**
     * @param array $element
     * @param string $attr
     * @return string|null
     */
    protected static function getValue(array $element, string $type = self::ATTR_VALUE): ?string
    {
        return $element['children'][$type][0]['text'] ?? null;
    }

    /**
     * get type by format
     * @param string $format
     * @return string
     */
    protected function getFormatTypes(string $format): string
    {
        if (preg_match('/(d{1,2}[ -.\/]m{1,3}[ -.\/]y{1,4})( h{1,2}:m{1,2})?/', $format, $matches)) {
            return count($matches) > 2
                ? static::TYPE_DATETIME
                : static::TYPE_DATE;
        }

        return static::TYPE_OTHER;
    }

    /**
     * Converts XML objects to an array
     * Function from http://php.net/manual/pt_BR/book.simplexml.php
     *
     * @param $node
     * @return array
     */
    protected function xmlObjToArray(SimpleXMLElement $node): array
    {
        $namespace = (array)$node->getDocNamespaces(true);
        $namespace[null] = null;

        $children = [];
        $attributes = [];

        $text = trim((string)$node);
        if (strlen($text) <= 0) {
            $text = null;
        }


        foreach ($namespace as $ns => $nsUrl) {
            $objAttributes = $node->attributes($ns, true);
            foreach ($objAttributes as $attributeName => $attributeValue) {
                $attributeName = strtolower(trim((string)$attributeName));
                $attributeValue = trim((string)$attributeValue);
                if (!empty($ns)) {
                    $attributeName = sprintf('%s:%s', $ns, $attributeName);
                }
                $attributes[$attributeName] = $attributeValue;
            }

            // Children
            $objChildren = $node->children($ns, true);
            foreach ($objChildren as $childName => $child) {
                $childName = strtolower((string)$childName);
                if (!empty($ns)) {
                    $childName = sprintf('%s:%s', $ns, $childName);
                }
                $children[$childName][] = $this->xmlObjToArray($child);
            }
        }

        return [
            'text' => $text,
            'attributes' => $attributes,
            'children' => $children,
        ];
    }

    /**
     * Write array to CSV file
     * Enhanced fputcsv found at http://php.net/manual/en/function.fputcsv.php
     *
     * @param $handle
     * @param array $fields
     */
    protected function writeArrayToCsv($handle, array $fields): void
    {
        $first = 1;
        foreach ($fields as $field) {
            if ($first === 0) {
                fwrite($handle, static::$delimeter);
            }

            $fixedField = str_replace(static::$enclosure, static::$enclosure . static::$enclosure, $field);
            if (static::$enclosure !== static::$escape) {
                $fixedField = str_replace(static::$escape . static::$enclosure, static::$escape, $fixedField);
            }

            if (strpbrk($fixedField, " \t\n\r" . static::$delimeter . static::$enclosure . static::$escape) || strpos($fixedField, "\000") !== false) {
                fwrite($handle, static::$enclosure . $fixedField . static::$enclosure);
            } else {
                fwrite($handle, $fixedField);
            }

            $first = 0;
        }

        fwrite($handle, "\n");
    }

    /**
     * @param $dir
     */
    protected function cleanUp($dir): void
    {
        $tempdir = opendir($dir);
        while (($file = readdir($tempdir)) !== false) {
            if ($file !== '.' && $file !== '..') {
                $path = "$dir/$file";
                if (is_dir($path)) {
                    chdir('.');
                    $this->cleanUp($path);
                    rmdir($path);
                } else {
                    unlink($path);
                }
            }
        }

        closedir($tempdir);
    }

    /**
     * @return bool
     * @throws XlsxConverterException
     */
    protected function unpack(): bool
    {
        if ($this->extracted) {
            return true;
        }

        if (!file_exists($this->xlsxPath)) {
            throw new XlsxConverterException('File not found');
        }

        $this->mkdir($this->tmpDir);

        $zip = new ZipArchive;

        if ($zip->open($this->xlsxPath) === true) {
            $zip->extractTo($this->tmpDir);
            $zip->close();
            $this->extracted = true;
            return true;
        }

        return false;
    }

    /**
     * @param string $dir
     * @return void
     * @throws XlsxConverterException
     */
    protected function mkdir(string $dir): void
    {
        if (is_dir($this->tmpDir)) {
            return;
        }

        if (!mkdir($dir, 0755) && !is_dir($dir)) {
            throw new XlsxConverterException(sprintf('Directory "%s" was not created', $dir));
        }
    }
}

<?php

namespace TmeApp\Services\Xlsx;

use Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;

class SpreadsheetService
{
    /**
     * Default path of json data for parse.
     */
    private const DEFAULT_PRODUCTS_DATA_FILE_PATH = __DIR__ . '/../../../data.json';
    private const DEFAULT_XLSX_FILE_PATH = __DIR__ . '/../../../var/tmp/products_data.xlsx';

    /**
     * Table with all xlsx file columns
     */
    private const HEADERS = [
        'MPN',
        'Stock',
        'Manufacturer',
        'URL',
        'Description',
        'Parameters',
        'Document',
        'Categories',
        'Unit'
    ];

    /**
     * Table with all xlsx file columns with their widths
     */
    private const COLUMNS_WIDTHS = [
        'MPN' => 20,
        'Stock' => 7,
        'Manufacturer' => 25,
        'URL' => 20,
        'Description' => 45,
        'Parameters' => 30,
        'Document' => 20,
        'Categories' => 25,
        'Unit' => 4,
    ];


    private const FIRST_COLUMN = 'A';
    private const HEADER_ROW = 1;
    private const FIRST_DATA_ROW = self::HEADER_ROW + 1;

    /**
     * Array for products data from decoded json file
     */
    private array $productsData;

    /**
     * Spreadsheet property for handling xlsx file.
     *
     * @var Spreadsheet
     */
    private Spreadsheet $spreadsheet;

    private int $rowsCount;
    private int $columnsCount;

    /**
     * Last column letter
     */
    private $lastColumn;

    /**
     * @throws Exception
     */
    public function __construct(string $dataFilePath = self::DEFAULT_PRODUCTS_DATA_FILE_PATH)
    {

        if (!file_exists($dataFilePath)) {
            throw new Exception('Plik z danymi nie istnieje.');
        }

        try {
            $file = file_get_contents($dataFilePath);
            $this->productsData = json_decode($file, true);
        } catch (Exception $e) {
            throw new Exception('Wystąpił błąd podczas przetwarzania danych z podanego pliku');
        }

        $this->rowsCount = count($this->productsData) + 1;
        $this->columnsCount = count(self::HEADERS);
        $this->lastColumn = $this->getColumnLetter($this->columnsCount);

        $this->spreadsheet = new Spreadsheet();

    }

    /**
     * Create and save products data xlsx file in tmp location.
     *
     * @return void
     * @throws Exception
     */
    public function createXlsx()
    {
        try {
            $this->prepareXlsx();
            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet);
            $writer->save(self::DEFAULT_XLSX_FILE_PATH);
        } catch (Exception $e) {
            throw new Exception('Wystąpił błąd podczas zapisywania pliku xls. Błąd: ' . $e->getMessage());
        }

    }

    /**
     * Perform all actions necessary for creating xlsx file
     *
     * @return void
     * @throws Exception
     */
    private function prepareXlsx()
    {
        try {
            $this->setHeaders();
            $this->setDataRows();
            $this->addStyling();
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }

    }


    /**
     * Set header columns for xlsx file
     *
     * @return void
     */
    private function setHeaders()
    {
        $columnLetter = self::FIRST_COLUMN;
        foreach (self::HEADERS as $header) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . self::HEADER_ROW, $header);
            $columnLetter++;
        }
    }

    /**
     * Sets products data rows in xlsx file
     *
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function setDataRows()
    {
        $rowCounter = self::FIRST_DATA_ROW;

        foreach ($this->productsData as $productData) {
            if (!is_array($productData)) {
                continue;
            }

            $this->spreadsheet = DataRowService::setSingleDataRow($this->spreadsheet, $productData, $rowCounter);
            $rowCounter++;
        }
    }

    /**
     * Perform all styling actions for new xlsx file
     *
     * @return void
     */
    private function addStyling()
    {
        $this->setColumnsWidths();
        $this->setHeadersBold();
        $this->setBorders();
        $this->setAlignment();
    }

    /**
     *
     * @return void
     */
    private function setColumnsWidths()
    {
        $columnIterator = self::FIRST_COLUMN;
        foreach (self::COLUMNS_WIDTHS as $name => $width) {
            $this->spreadsheet->getActiveSheet()->getColumnDimension($columnIterator)->setWidth($width);
            $columnIterator++;
        }
    }

    /**
     * @return void
     */
    private function setHeadersBold()
    {

        $this->spreadsheet->getActiveSheet()->getStyle(self::FIRST_COLUMN . self::HEADER_ROW . ':' . $this->getColumnLetter(count(self::HEADERS)) . self::HEADER_ROW)
            ->getFont()->setBold('true');

    }

    /**
     * Sets oll borders for new xlsx file
     *
     * @return void
     */
    private function setBorders()
    {

        for ($row = 1; $row <= $this->rowsCount; $row++) {

            if ($row == 1) {

                for ($column = 'A'; $column <= $this->lastColumn; $column++) {

                    $this->spreadsheet->getActiveSheet()->getStyle($column . $row)
                        ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
                    $this->spreadsheet->getActiveSheet()->getStyle($column . $row)
                        ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
                    $this->spreadsheet->getActiveSheet()->getStyle($column . $row)
                        ->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
                    $this->spreadsheet->getActiveSheet()->getStyle($column . $row)
                        ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
                }

            } else {
                $this->spreadsheet->getActiveSheet()->getStyle(self::FIRST_COLUMN . $row . ':' . $this->lastColumn . $this->rowsCount)
                    ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                $this->spreadsheet->getActiveSheet()->getStyle(self::FIRST_COLUMN . $row . ':' . $this->lastColumn . $this->rowsCount)
                    ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                $this->spreadsheet->getActiveSheet()->getStyle($this->lastColumn . $row)
                    ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            }

        }
    }

    /**
     * Get column letter by given column number
     *
     * @param int $columnNumber
     * @return mixed
     */
    private function getColumnLetter(int $columnNumber)
    {
        $alphabet = range('A', 'Z');
        return $alphabet[$columnNumber - 1];
    }

    /**
     * Sets alignment for data inside xlsx file.
     *
     * @return void
     */
    private function setAlignment()
    {
        $this->spreadsheet->getActiveSheet()->getStyle(self::FIRST_COLUMN . self::FIRST_DATA_ROW . ':' . $this->lastColumn . $this->rowsCount)
            ->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

        $this->spreadsheet->getActiveSheet()->getStyle(self::FIRST_COLUMN . self::FIRST_DATA_ROW . ':' . $this->lastColumn . $this->rowsCount)
            ->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
    }

    /**
     * @return string
     */
    public static function getDefaultXlsxFilePath(): string
    {
        return self::DEFAULT_XLSX_FILE_PATH;
    }
}

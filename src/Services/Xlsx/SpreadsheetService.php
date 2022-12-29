<?php

namespace TmeApp\Services\Xlsx;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;

class SpreadsheetService
{
    private const DEFAULT_DATA_FILE_PATH = __DIR__ . '/../../../data.json';

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

    const FIRST_COLUMN = 'A';
    const HEADER_ROW = 1;
    const FIRST_DATA_ROW = self::HEADER_ROW + 1;

    private array $productsData;

    private Spreadsheet $spreadsheet;

    private $rowsCount;
    private $columnsCount;
    private $lastColumn;

    public function __construct(string $dataFilePath = self::DEFAULT_DATA_FILE_PATH)
    {

        if(!file_exists($dataFilePath)){
            throw new \Exception('Plik z danymi nie istnieje.');
        }

        try {
            $file = file_get_contents($dataFilePath);
            $this->productsData = json_decode($file,true);
        }catch (\Exception $e){
            throw new \Exception('Wystąpił błąd podczas przetwarzania danych z podanego pliku');
        }

        $this->rowsCount = count($this->productsData) + 1;
        $this->columnsCount = count(self::HEADERS);
        $this->lastColumn = $this->getColumnLetter($this->columnsCount);

        $this->spreadsheet = new Spreadsheet();

    }

    public function getXlsx()
    {

        $this->prepareXlsx();
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet);
        try {
            $writer->save(__DIR__ . "/../../../var/tmp/products_data.xlsx");
        } catch (\Exception $e) {
            throw new \Exception('Wystąpił błąd podczas zapisywania pliku xls. Błąd: ' . $e->getMessage());
        }

    }

    private function prepareXlsx()
    {
        $this->setHeaders();
        $this->setDataRows();
        $this->addStyling();
    }



    private function setHeaders()
    {
        $columnLetter = self::FIRST_COLUMN;
        foreach (self::HEADERS as $header){
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter .self::HEADER_ROW,$header);
            $columnLetter++;
        }
    }

    private function setDataRows()
    {
        $rowCounter = self::FIRST_DATA_ROW;

        foreach ($this->productsData as $productData){
            if(!is_array($productData)){
                continue;
            }

            $this->spreadsheet = DataRowService::setSingleDataRow($this->spreadsheet,$productData,$rowCounter);
            $rowCounter++;
        }
    }

    private function addStyling()
    {
        $this->setColumnsWidths();
        $this->setHeadersColors();
        $this->setBorders();
        $this->setAlignment();
    }

    private function setColumnsWidths()
    {
        $columnIterator = self::FIRST_COLUMN;
        foreach (self::COLUMNS_WIDTHS as $name => $width) {
            $this->spreadsheet->getActiveSheet()->getColumnDimension($columnIterator)->setWidth($width);
            $columnIterator++;
        }
    }

    private function setHeadersColors(){

                $this->spreadsheet->getActiveSheet()->getStyle(self::FIRST_COLUMN .  self::HEADER_ROW . ':' . $this->getColumnLetter(count(self::HEADERS)) . self::HEADER_ROW)
                    ->getFont()->setBold('true');

        }

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

    private function getColumnLetter(int $columnNumber){
        $alphabet = range('A', 'Z');
        return $alphabet[$columnNumber - 1];
    }

    private function setAlignment()
    {
        $this->spreadsheet->getActiveSheet()->getStyle(self::FIRST_COLUMN . self::FIRST_DATA_ROW . ':' . $this->lastColumn . $this->rowsCount)
            ->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

        $this->spreadsheet->getActiveSheet()->getStyle(self::FIRST_COLUMN . self::FIRST_DATA_ROW . ':' . $this->lastColumn . $this->rowsCount)
            ->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
    }

}

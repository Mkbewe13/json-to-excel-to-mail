<?php

namespace TmeApp\Services\Xlsx;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

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

    const FIRST_COLUMN = 'A';
    const HEADER_ROW = 1;
    const FIRST_DATA_ROW = self::HEADER_ROW + 1;

    private array $productsData;

    private Spreadsheet $spreadsheet;

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

        $this->spreadsheet = new Spreadsheet();

    }

    public function getXlsx(){

        $this->prepareXlsx();
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet);
        $writer->save("demo.xlsx");
    }

    private function prepareXlsx()
    {
        $this->setHeaders();
        $this->setDataRows();
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



}

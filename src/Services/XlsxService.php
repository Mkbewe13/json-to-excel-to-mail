<?php

namespace TmeApp\Services;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

class XlsxService
{
    private const DEFAULT_DATA_FILE_PATH = __DIR__.'/../../data.json';

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

            $this->setSingleDataRow($productData,$rowCounter);
            $rowCounter++;
        }
    }

    private function setSingleDataRow(array $productData,int $rowNumber)
    {
        $this->setMPN($productData['mpn'] ?? null,$rowNumber);
        $this->setStock($productData['stock'] ?? null,$rowNumber);
        $this->setManufacturer($productData['manufacturer'] ?? null,$rowNumber);
        $this->setURL($productData['url'] ?? null,$rowNumber);
        $this->setDescription($productData['description'] ?? null,$rowNumber);
        $this->setParameters($productData['parametersAsString'] ?? null,$rowNumber);
        $this->setDocument($productData['documents'] ?? null,$rowNumber);
        $this->setCategories($productData['breadcrumbs'] ?? null,$rowNumber);
        $this->setUnit($productData['unit'] ?? null,$rowNumber);
    }

    private function setMPN(?string $mpn, int $rowNumber)
    {
        $columnLetter = 'A';

        if ($mpn === null) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }
        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $mpn);
    }

    private function setStock(?int $stock, int $rowNumber)
    {
        $columnLetter = 'B';

        if ($stock === null) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }
        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $stock);
    }

    private function setManufacturer(?array $manufacturer, int $rowNumber)
    {
        $columnLetter = 'C';
        if (empty($manufacturer) || empty($manufacturer['name'])) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }

        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $manufacturer['name']);
    }

    private function setURL(?string $url, int $rowNumber)
    {
        $columnLetter = 'D';

        if ($url === null) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }

        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $url);
    }

    private function setDescription(?string $description, int $rowNumber)
    {
        $columnLetter = 'E';

        if ($description === null) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }

        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $description);
    }

    private function setParameters(?string $parameters, int $rowNumber)
    {
        $columnLetter = 'F';

        if ($parameters === null) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }

        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $parameters);
    }

    private function setDocument(array $documents, int $rowNumber)
    {

        $columnLetter = 'G';

        if (empty($documents)) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }

        $englishDocumentUrl = 'b/d';
        foreach ($documents as $document){
            if(empty($document['type']) || empty($document['language']) || empty($document['url'])){
                continue;
            }

            if($document['type'] === 'DTE' && $document['language'] === 'PL'){
                $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $document['url']);
                return;
            }

            if($document['type'] === 'DTE' && $document['language'] === 'EN' && $englishDocumentUrl === 'b/d'){
                $englishDocumentUrl = $document['url'];
            }
        }

        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $englishDocumentUrl);
    }

    private function setUnit(?array $unit, int $rowNumber)
    {
        $columnLetter = 'I';

        if (empty($unit) || empty($unit['label'])) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }


        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $unit['label']);
    }

    private function setCategories(?string $breadcrumbs, int $rowNumber)
    {
        $columnLetter = 'H';

        if ($breadcrumbs === null) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }


        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $breadcrumbs);
    }


}

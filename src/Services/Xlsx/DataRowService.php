<?php

namespace TmeApp\Services\Xlsx;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;

class DataRowService
{
    private Spreadsheet $spreadsheet;

    private function __construct($spreadsheet)
    {
        $this->spreadsheet = $spreadsheet;
    }

    public static function setSingleDataRow(Spreadsheet $spreadsheet,array $productData,int $rowNumber): Spreadsheet
    {
        $dataRowService = new self($spreadsheet);

        $dataRowService->setMPN($productData['mpn'] ?? null,$rowNumber);
        $dataRowService->setStock($productData['stock'] ?? null,$rowNumber);
        $dataRowService->setManufacturer($productData['manufacturer'] ?? null,$rowNumber);
        $dataRowService->setURL($productData['url'] ?? null,$rowNumber);
        $dataRowService->setDescription($productData['description'] ?? null,$rowNumber);
        $dataRowService->setParameters($productData['parametersAsString'] ?? null,$rowNumber);
        $dataRowService->setDocument($productData['documents'] ?? null,$rowNumber);
        $dataRowService->setCategories($productData['breadcrumbs'] ?? null,$rowNumber);
        $dataRowService->setUnit($productData['unit'] ?? null,$rowNumber);

        return $dataRowService->spreadsheet;
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
        $this->spreadsheet->getActiveSheet()->getCell($columnLetter . $rowNumber)->getHyperlink()->setUrl($url);
        $this->spreadsheet->getActiveSheet()->getStyle($columnLetter . $rowNumber)
            ->getFont()->getColor()->setARGB(Color::COLOR_BLUE);

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

        if ($parameters === null || $parameters === '') {
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
                $this->spreadsheet->getActiveSheet()->getCell($columnLetter . $rowNumber)->getHyperlink()->setUrl($document['url']);
                $this->spreadsheet->getActiveSheet()->getStyle($columnLetter . $rowNumber)
                    ->getFont()->getColor()->setARGB(Color::COLOR_BLUE);
                return;
            }

            if($document['type'] === 'DTE' && $document['language'] === 'EN' && $englishDocumentUrl === 'b/d'){
                $englishDocumentUrl = $document['url'];
            }
        }

        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $englishDocumentUrl);
        $this->spreadsheet->getActiveSheet()->getCell($columnLetter . $rowNumber)->getHyperlink()->setUrl($englishDocumentUrl);
        $this->spreadsheet->getActiveSheet()->getStyle($columnLetter . $rowNumber)
            ->getFont()->getColor()->setARGB(Color::COLOR_BLUE);
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

<?php

namespace TmeApp\Services\Xlsx;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;

/**
 * Class for handling single row of spreadsheet data.
 */
class DataRowService
{
    private Spreadsheet $spreadsheet;

    private function __construct($spreadsheet)
    {
        $this->spreadsheet = $spreadsheet;
    }

    /**
     * Sets single data row for given spreadsheet by productData array and rownumber.
     * Returns spreadsheet object
     *
     * @param Spreadsheet $spreadsheet
     * @param array $productData
     * @param int $rowNumber
     * @return Spreadsheet
     * @throws Exception
     * @throws \Exception
     */
    public static function setSingleDataRow(Spreadsheet $spreadsheet, array $productData, int $rowNumber): Spreadsheet
    {
        try {
            $dataRowService = new self($spreadsheet);

            $dataRowService->setMPN($productData['mpn'] ?? null, $rowNumber);
            $dataRowService->setStock($productData['stock'] ?? null, $rowNumber);
            $dataRowService->setManufacturer($productData['manufacturer'] ?? null, $rowNumber);
            $dataRowService->setURL($productData['url'] ?? null, $rowNumber);
            $dataRowService->setDescription($productData['description'] ?? null, $rowNumber);
            $dataRowService->setParameters($productData['parametersAsString'] ?? null, $rowNumber);
            $dataRowService->setDocument($productData['documents'] ?? null, $rowNumber);
            $dataRowService->setCategories($productData['breadcrumbs'] ?? null, $rowNumber);
            $dataRowService->setUnit($productData['unit'] ?? null, $rowNumber);

            return $dataRowService->spreadsheet;
        } catch (\Exception $e) {
            throw new \Exception($e->getMessage());
        }

    }

    /**
     * @param string|null $mpn
     * @param int $rowNumber
     * @return void
     */
    private function setMPN(?string $mpn, int $rowNumber)
    {
        $columnLetter = 'A';

        if ($mpn === null) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }
        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $mpn);
    }

    /**
     * @param int|null $stock
     * @param int $rowNumber
     * @return void
     */
    private function setStock(?int $stock, int $rowNumber)
    {
        $columnLetter = 'B';

        if ($stock === null) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }
        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $stock);
    }

    /**
     * @param array|null $manufacturer
     * @param int $rowNumber
     * @return void
     */
    private function setManufacturer(?array $manufacturer, int $rowNumber)
    {
        $columnLetter = 'C';
        if (empty($manufacturer) || empty($manufacturer['name'])) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }

        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $manufacturer['name']);
    }

    /**
     * @param string|null $url
     * @param int $rowNumber
     * @return void
     * @throws Exception
     */
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

    /**
     * @param string|null $description
     * @param int $rowNumber
     * @return void
     */
    private function setDescription(?string $description, int $rowNumber)
    {
        $columnLetter = 'E';

        if ($description === null) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }

        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $description);
    }

    /**
     * @param string|null $parameters
     * @param int $rowNumber
     * @return void
     */
    private function setParameters(?string $parameters, int $rowNumber)
    {
        $columnLetter = 'F';

        if ($parameters === null || $parameters === '') {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }


        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $parameters);
    }

    /**
     * @param array $documents
     * @param int $rowNumber
     * @return void
     * @throws Exception
     */
    private function setDocument(array $documents, int $rowNumber)
    {

        $columnLetter = 'G';

        if (empty($documents)) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }

        $englishDocumentUrl = 'b/d';
        foreach ($documents as $document) {
            if (empty($document['type']) || empty($document['language']) || empty($document['url'])) {
                continue;
            }

            if ($document['type'] === 'DTE' && $document['language'] === 'PL') {
                $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $document['url']);
                $this->spreadsheet->getActiveSheet()->getCell($columnLetter . $rowNumber)->getHyperlink()->setUrl($document['url']);
                $this->spreadsheet->getActiveSheet()->getStyle($columnLetter . $rowNumber)
                    ->getFont()->getColor()->setARGB(Color::COLOR_BLUE);
                return;
            }

            if ($document['type'] === 'DTE' && $document['language'] === 'EN' && $englishDocumentUrl === 'b/d') {
                $englishDocumentUrl = $document['url'];
            }
        }

        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $englishDocumentUrl);
        $this->spreadsheet->getActiveSheet()->getCell($columnLetter . $rowNumber)->getHyperlink()->setUrl($englishDocumentUrl);
        $this->spreadsheet->getActiveSheet()->getStyle($columnLetter . $rowNumber)
            ->getFont()->getColor()->setARGB(Color::COLOR_BLUE);
    }

    /**
     * @param array|null $unit
     * @param int $rowNumber
     * @return void
     */
    private function setUnit(?array $unit, int $rowNumber)
    {
        $columnLetter = 'I';

        if (empty($unit) || empty($unit['label'])) {
            $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, 'b/d');
            return;
        }


        $this->spreadsheet->getActiveSheet()->setCellValue($columnLetter . $rowNumber, $unit['label']);
    }

    /**
     * @param string|null $breadcrumbs
     * @param int $rowNumber
     * @return void
     */
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

<?php

/**
 * Create completed xlsx for bar/cols, line, radar,... charts
 *
 * @category   Phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateCompletedXlsx extends CreateXlsx
{
    /**
     * Create excel sheet
     *
     * @access public
     * @param array $dats
     */
    public function createExcelSheet($dats)
    {
        $this->xml = '';
        $sizeDats = count($dats['data']);
        $sizeCols = count($dats['data'][0]['values']);
        $this->generateWORKSHEET();
        $this->generateDIMENSION($sizeDats, $sizeCols);
        $this->generateSHEETVIEWS();
        $this->generateSHEETVIEW();
        $this->generateSELECTION($sizeDats + $sizeCols);
        $this->generateSHEETFORMATPR();
        $this->generateCOLS();
        $this->generateCOL();
        $this->generateSHEETDATA();
        $row = 1;

        // add the legend
        $col = 1;
        $letter = 'A';
        $this->generateROW($row, $sizeCols);
        $this->generateC($letter . $row, '', 's');
        $this->generateV($sizeDats + $sizeCols);
        if (function_exists('str_increment')) {
            $letter = str_increment($letter);
        } else {
            $letter++;
        }
        if (isset($dats['legend'])) {
            foreach ($dats['legend'] as $legend) {
                $this->generateC($letter . $row, '', 's');
                $this->generateV($col - 1);
                $col++;
                if (function_exists('str_increment')) {
                    $letter = str_increment($letter);
                } else {
                    $letter++;
                }
            }

            $this->cleanTemplateROW();
            $row++;
        } else {
            $this->generateC($letter . $row, '', 's');
            $this->generateV($col - 1);
            $col++;
            if (function_exists('str_increment')) {
                $letter = str_increment($letter);
            } else {
                $letter++;
            }
            $this->cleanTemplateROW();
            $row++;
        }

        // add the data
        foreach ($dats['data'] as $data) {
            $this->generateROW($row, $sizeCols);
            $col = 1;
            $letter = 'A';
            $this->generateC($letter . $row, 1, 's');
            $this->generateV($sizeCols + $row - 2);
            if (function_exists('str_increment')) {
                $letter = str_increment($letter);
            } else {
                $letter++;
            }
            foreach ($data['values'] as $value) {
                $s = '';
                if ($col != $sizeCols) {
                    $s = 1;
                }
                $this->generateC($letter . $row, $s);
                $this->generateV($value);
                $col++;
                if (function_exists('str_increment')) {
                    $letter = str_increment($letter);
                } else {
                    $letter++;
                }
            }
            $row++;
            $this->cleanTemplateROW();
        }
        $this->generateROW($row + 1, $sizeCols);
        $row++;
        $this->generateC('B' . $row, 2, 's');
        $this->generateV($sizeDats + $sizeCols + 1);
        $this->generatePAGEMARGINS();
        $this->generateTABLEPARTS();
        $this->generateTABLEPART(1);
        $this->cleanTemplate();

        return $this->xml;
    }

    /**
     * Create excel shared strings
     *
     * @access public
     * @param array $dats
     */
    public function createExcelSharedStrings($dats)
    {
        $this->xml = '';
        $szDats = count($dats['data']);
        $szCols = count($dats['data'][0]['values']);
        $this->generateSST($szDats + $szCols + 2);

        for ($i = 0; $i < $szCols; $i++) {
            $this->generateSI();
            if (!isset($dats['legend'][$i])) {
                $dats['legend'][$i] = '0';
            }
            $this->generateT($dats['legend'][$i]);
        }

        foreach ($dats['data'] as $data) {
            $this->generateSI();
            if (isset($data['name'])) {
                $this->generateT($data['name']);
            }
        }
        $this->generateSI();
        $this->generateT(' ', 'preserve');

        $msg = 'To change the range size of values, drag the bottom right corner';
        $this->generateSI();
        $this->generateT($msg);

        $this->cleanTemplate();

        return $this->xml;
    }

    /**
     * Create excel table
     *
     * @access public
     * @param array $dats
     */
    public function createExcelTable($dats)
    {
        $this->xml = '';
        $szDats = count($dats['data']);
        $szCols = count($dats['data'][0]['values']);
        $this->generateTABLE($szDats, $szCols);
        $this->generateTABLECOLUMNS($szCols + 1);
        $this->generateTABLECOLUMN(1, ' ');
        for ($i = 0; $i < $szCols; $i++) {
            if (!isset($dats['legend'][$i])) {
                $dats['legend'][$i] = '0';
            }
            $this->generateTABLECOLUMN($i + 2, $dats['legend'][$i]);
        }
        $this->generateTABLESTYLEINFO();
        $this->cleanTemplate();

        return $this->xml;
    }
}
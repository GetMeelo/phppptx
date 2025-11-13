<?php

/**
 * Create xlsx for scatter charts
 *
 * @category   phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateScatterXlsx extends CreateXlsx
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
        $sizeCols = 2;
        $this->generateWORKSHEET();
        $this->generateDIMENSION($sizeDats + 1, $sizeCols);
        $this->generateSHEETVIEWS();
        $this->generateSHEETVIEW();
        $this->generateSELECTION($sizeDats + $sizeCols - 1);
        $this->generateSHEETFORMATPR();
        $this->generateCOLS();
        $this->generateCOL();
        $this->generateSHEETDATA();
        $row = 1;
        $col = 1;
        $letter = 'A';
        $this->generateROW($row, $sizeCols - 1);
        for ($num = 0; $num < $sizeCols; $num++) {
            $this->generateC($letter . $row, '1', 's');
            $this->generateV($num);
            if (function_exists('str_increment')) {
                $letter = str_increment($letter);
            } else {
                $letter++;
            }
            $col++;
        }
        $this->cleanTemplateROW();
        $row++;

        foreach ($dats['data'] as $data) {
            $this->generateROW($row, $sizeCols - 1);
            $col = 1;
            $letter = 'A';
            foreach ($data['values'] as $values) {
                $this->generateC($letter . $row, '', 'n');
                if (is_array($values)) {
                    foreach ($values as $valuesInternalArray) {
                        $this->generateV($valuesInternalArray);
                    }
                } else {
                    $this->generateV($values);
                }
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
        $this->generateV($num + 1);
        $this->generatePAGEMARGINS();
        $this->generateTABLEPARTS();
        $this->generateTABLEPART(1);
        $this->cleanTemplate();

        return $this->xml;
    }

    /**
     * Create excel shared strings
     *
     * @param array $dats
     * @access public
     */
    public function createExcelSharedStrings($dats)
    {
        $this->xml = '';
        $szDats = count($dats['data']);
        $szCols = 2;
        $this->generateSST($szCols + 2);
        $legends = array('X Values', 'Y Values');
        for ($i = 0; $i < $szCols; $i++) {
            $this->generateSI();
            if (isset($dats['legend']) && isset($dats['legend'][$i])) {
                $legends[$i] = $dats['legend'][$i];
            }
            $this->generateT($legends[$i]);
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
        $szCols = 2;
        $this->generateTABLE($szDats, $szCols - 1);
        $this->generateTABLECOLUMNS($szCols);
        $legends = array('X Values', 'Y Values');
        for ($i = 0; $i < $szCols; $i++) {
            if (isset($dats['legend']) && isset($dats['legend'][$i])) {
                $legends[$i] = $dats['legend'][$i];
            }
            $this->generateTABLECOLUMN($i + 1, $legends[$i]);
        }
        $this->generateTABLESTYLEINFO();
        $this->cleanTemplate();

        return $this->xml;
    }

    /**
     * Generate dimension
     *
     * @access protected
     * @param int $sizeX
     * @param int $sizeY
     */
    protected function generateDIMENSION($sizeX, $sizeY)
    {
        $char = 'A';
        for ($i = 0; $i < $sizeY - 1; $i++) {
            if (function_exists('str_increment')) {
                $char = str_increment($char);
            } else {
                $char++;
            }
        }
        $sizeX += $sizeY;
        $xml = '<dimension ref="A1:' . $char . $sizeX . '"></dimension>__PHX=__GENERATEWORKSHEET__';

        $this->xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->xml);
    }
}
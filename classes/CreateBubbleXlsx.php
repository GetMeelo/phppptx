<?php

/**
 * Create xlsx for bubble charts
 *
 * @category   phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateBubbleXlsx extends CreateXlsx
{
    /**
     * Create excel sheet
     *
     * @access public
     * @param array $dats
     */
    public function createExcelSheet($dats)
    {
        unset($dats['legend']); // eliminate this row because of the way bubble charts handle the excel data
        $this->xml = '';
        $sizeDats = count($dats['data']);
        $sizeCols = 3;
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
        $col = 1;
        $letter = 'A';
        $this->generateROW($row, $sizeCols - 1);
        for ($num = 0; $num < $sizeCols; $num++) {
            $this->generateC($letter . $row, '', 's');
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
                $this->generateV($values);
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
     * @access public
     * @param array $dats
     */
    public function createExcelSharedStrings($dats)
    {
        $this->xml = '';
        $szDats = count($dats['data']);
        $szCols = 3;
        $this->generateSST(($szDats + 1) * 3 + 2);
        $legends = array('X-Values', 'Y-Values', 'Size');
        if (!empty($dats['legend'][0])) {
            $legends[0] = $dats['legend'][0];
        }
        if (!empty($dats['legend'][1])) {
            $legends[1] = $dats['legend'][1];
        }
        if (!empty($dats['legend'][2])) {
            $legends[2] = $dats['legend'][2];
        }
        for ($i = 0; $i < $szCols; $i++) {
            $this->generateSI();
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
        $szCols = 3;
        $this->generateTABLE($szDats, $szCols - 1);
        $this->generateTABLECOLUMNS($szCols);
        $legends = array('X-Values', 'Y-Values', 'Size');
        if (!empty($dats['legend'][0])) {
            $legends[0] = $dats['legend'][0];
        }
        if (!empty($dats['legend'][1])) {
            $legends[1] = $dats['legend'][1];
        }
        if (!empty($dats['legend'][2])) {
            $legends[2] = $dats['legend'][2];
        }
        for ($i = 0; $i < $szCols; $i++) {
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
        //$sizeY--;//to get rid of the legends row
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
<?php

/**
 * Create XLSX
 *
 * @category   Phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateXlsx extends CreateElement
{
    /**
     *
     * @access private
     * @var PptxStructure
     */
    private $zipXlsx;

    // empty functions to be overridden

    /**
     * Create excel table
     *
     * @access public
     * @param array $dats
     */
    public function createExcelTable($dats) {}
    /**
     * Create excel shared strings
     *
     * @access public
     * @param array $dats
     */
    public function createExcelSharedStrings($dats) {}
    /**
     * Create excel sheet
     *
     * @access public
     * @param array $dats
     */
    public function createExcelSheet($dats) {}

    /**
     * Create XLSX
     *
     * @access public
     * @param array $idChart
     * @param array $chartData
     * @return PptxStructure
     */
    public function createChartXlsx($idChart, $chartData)
    {
        $this->zipXlsx = new PptxStructure();

        $this->zipXlsx->addContent('[Content_Types].xml', ExcelStructureTemplate::$excelStructure['[Content_Types].xml']);
        $this->zipXlsx->addContent('docProps/core.xml', ExcelStructureTemplate::$excelStructure['docProps/core.xml']);
        $this->zipXlsx->addContent('docProps/app.xml', ExcelStructureTemplate::$excelStructure['docProps/app.xml']);
        $this->zipXlsx->addContent('_rels/.rels', ExcelStructureTemplate::$excelStructure['_rels/.rels']);
        $this->zipXlsx->addContent('xl/_rels/workbook.xml.rels', ExcelStructureTemplate::$excelStructure['xl/_rels/workbook.xml.rels']);
        $this->zipXlsx->addContent('xl/theme/theme1.xml', ExcelStructureTemplate::$excelStructure['xl/theme/theme1.xml']);
        $this->zipXlsx->addContent('xl/worksheets/_rels/sheet1.xml.rels', ExcelStructureTemplate::$excelStructure['xl/worksheets/_rels/sheet1.xml.rels']);
        $this->zipXlsx->addContent('xl/styles.xml', ExcelStructureTemplate::$excelStructure['xl/styles.xml']);
        $this->zipXlsx->addContent('xl/workbook.xml', ExcelStructureTemplate::$excelStructure['xl/workbook.xml']);
        $this->zipXlsx->addContent('xl/tables/table1.xml', $this->createExcelTable($chartData));
        $this->zipXlsx->addContent('xl/sharedStrings.xml', $this->createExcelSharedStrings($chartData));
        $this->zipXlsx->addContent('xl/worksheets/sheet1.xml', $this->createExcelSheet($chartData));

        return $this->zipXlsx;
    }

    /**
     * Generate sst
     *
     * @access protected
     * @param mixed $num
     */
    protected function generateSST($num)
    {
        $this->xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' . $num . '" uniqueCount="' . $num . '">__PHX=__GENERATESST__</sst>';
    }

    /**
     * Generate si
     *
     * @access protected
     */
    protected function generateSI()
    {
        $xml = '<si>__PHX=__GENERATESI__</si>__PHX=__GENERATESST__';

        $this->xml = str_replace('__PHX=__GENERATESST__', $xml, $this->xml);
    }

    /**
     * Generate t
     *
     * @access protected
     * @param string $name
     * @param string $space
     */
    protected function generateT($name, $space = '')
    {
        $xmlAux = '<t';
        if ($space != '') {
            $xmlAux .= ' xml:space="' . $space . '"';
        }
        $xmlAux .= '>' . $this->parseAndCleanTextString($name) . '</t>';

        $this->xml = str_replace('__PHX=__GENERATESI__', $xmlAux, $this->xml);
    }

    /**
     * Generate c
     *
     * @access protected
     * @param mixed $r
     * @param mixed $s
     * @param mixed $t
     */
    protected function generateC($r, $s, $t = '')
    {
        $xmlAux = '<c r="' . $r . '"';
        if ($s != '') {
            $xmlAux .= ' s="' . $s . '"';
        }
        if ($t != '') {
            $xmlAux .= ' t="' . $t . '"';
        }
        $xmlAux .= '>__PHX=__GENERATEC__</c>__PHX=__GENERATEROW__';

        $this->xml = str_replace('__PHX=__GENERATEROW__', $xmlAux, $this->xml);
    }

    /**
     * Generate col
     *
     * @access protected
     * @param string $min
     * @param string $max
     * @param string $width
     * @param string $customWidth
     */
    protected function generateCOL($min = '1', $max = '1', $width = '11.85546875', $customWidth = '1')
    {
        $xml = '<col min="' . $min . '" max="' . $max . '" width="' . $width . '" customWidth="' . $customWidth . '"></col>';

        $this->xml = str_replace('__PHX=__GENERATECOLS__', $xml, $this->xml);
    }

    /**
     * Generate cols
     *
     * @access protected
     */
    protected function generateCOLS()
    {
        $xml = '<cols>__PHX=__GENERATECOLS__</cols>__PHX=__GENERATEWORKSHEET__';

        $this->xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->xml);
    }

    /**
     * Generate dimension
     *
     * @param int $sizeX
     * @param int $sizeY
     * @access protected
     */
    protected function generateDIMENSION($sizeX, $sizeY)
    {
        $char = 'A';
        for ($i = 0; $i < $sizeY; $i++) {
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

    /**
     * Generate pagemargins
     *
     * @param string $left
     * @param string $rigth
     * @param string $bottom
     * @param string $top
     * @param string $header
     * @param string $footer
     * @access protected
     */
    protected function generatePAGEMARGINS($left = '0.7', $rigth = '0.7', $bottom = '0.75', $top = '0.75', $header = '0.3', $footer = '0.3')
    {
        $xml = '<pageMargins left="' . $left . '" right="' . $rigth . '" top="' . $top . '" bottom="' . $bottom . '" header="' . $header . '" footer="' . $footer . '"></pageMargins>__PHX=__GENERATEWORKSHEET__';

        $this->xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->xml);
    }

    /**
     * Generate pagesetup
     *
     * @param string $paperSize
     * @param string $orientation
     * @param string $id
     * @access protected
     */
    protected function generatePAGESETUP($paperSize = '9', $orientation = 'portrait', $id = '1')
    {
        $xml = '<pageSetup paperSize="' . $paperSize . '" orientation="' . $orientation . '" r:id="rId' . $id . '"></pageSetup>__PHX=__GENERATEWORKSHEET__';

        $this->xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->xml);
    }

    /**
     * Generate row
     *
     * @access protected
     * @param mixed $r
     * @param mixed $spans
     */
    protected function generateROW($r, $spans)
    {
        $spans = '1:' . ($spans + 1);
        $xml = '<row r="' . $r . '" spans="' . $spans . '">__PHX=__GENERATEROW__</row>__PHX=__GENERATESHEETDATA__';

        $this->xml = str_replace('__PHX=__GENERATESHEETDATA__', $xml, $this->xml);
    }

    /**
     * Generate selection
     *
     * @access protected
     * @param mixed $num
     */
    protected function generateSELECTION($num)
    {
        $xml = '<selection activeCell="B' . $num . '" sqref="B' . $num . '"></selection>';

        $this->xml = str_replace('__PHX=__GENERATESHEETVIEW__', $xml, $this->xml);
    }

    /**
     * Generate sheetdata
     *
     * @access protected
     */
    protected function generateSHEETDATA()
    {
        $xml = '<sheetData>__PHX=__GENERATESHEETDATA__</sheetData>' . '__PHX=__GENERATEWORKSHEET__';

        $this->xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->xml);
    }

    /**
     * Generate sheetformatpr
     *
     * @param string $baseColWidth
     * @param string $defaultRowHeight
     * @access protected
     */
    protected function generateSHEETFORMATPR($baseColWidth = '10', $defaultRowHeight = '15')
    {
        $xml = '<sheetFormatPr baseColWidth="' . $baseColWidth . '" defaultRowHeight="' . $defaultRowHeight . '"></sheetFormatPr>__PHX=__GENERATEWORKSHEET__';

        $this->xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->xml);
    }

    /**
     * Generate sheetview
     *
     * @param string $tabSelected
     * @param string $workbookViewId
     * @access protected
     */
    protected function generateSHEETVIEW($tabSelected = '1', $workbookViewId = '0')
    {
        $xml = '<sheetView tabSelected="' . $tabSelected . '" workbookViewId="' . $workbookViewId . '">__PHX=__GENERATESHEETVIEW__</sheetView>';

        $this->xml = str_replace('__PHX=__GENERATESHEETVIEWS__', $xml, $this->xml);
    }

    /**
     * Generate sheetviews
     *
     * @access protected
     */
    protected function generateSHEETVIEWS()
    {
        $xml = '<sheetViews>__PHX=__GENERATESHEETVIEWS__</sheetViews>__PHX=__GENERATEWORKSHEET__';

        $this->xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->xml);
    }

    /**
     * Generate tablepart
     *
     * @access protected
     * @param mixed $id
     */
    protected function generateTABLEPART($id = '1')
    {
        $xml = '<tablePart r:id="rId' . $id . '"></tablePart>';

        $this->xml = str_replace('__PHX=__GENERATETABLEPARTS__', $xml, $this->xml);
    }

    /**
     * Generate tableparts
     *
     * @access protected
     * @param string $count
     */
    protected function generateTABLEPARTS($count = '1')
    {
        $xml = '<tableParts count="' . $count . '">__PHX=__GENERATETABLEPARTS__</tableParts>';

        $this->xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->xml);
    }

    /**
     * Generate v
     *
     * @access protected
     * @param mixed $num
     */
    protected function generateV($num)
    {
        $this->xml = str_replace('__PHX=__GENERATEC__', '<v>' . $num . '</v>', $this->xml);
    }

    /**
     * Generate worksheet
     *
     * @access protected
     */
    protected function generateWORKSHEET()
    {
        $this->xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' . 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">__PHX=__GENERATEWORKSHEET__</worksheet>';
    }

    /**
     * Clean template row tags
     *
     * @access private
     */
    protected function cleanTemplateROW()
    {
        $this->xml = str_replace('__PHX=__GENERATEROW__', '', $this->xml);
    }

    /**
     * Generate table
     *
     * @access protected
     * @param int $rows
     * @param int $cols
     */
    protected function generateTABLE($rows, $cols)
    {
        $word = 'A';
        for ($i = 0; $i < $cols; $i++) {
            if (function_exists('str_increment')) {
                $word = str_increment($word);
            } else {
                $word++;
            }
        }
        $rows++;
        $this->xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Tabla1" displayName="Tabla1" ref="A1:' . $word . $rows . '" totalsRowShown="0" tableBorderDxfId="0">__PHX=__GENERATETABLE__</table>';
    }

    /**
     * Generate tablecolumn
     *
     * @access protected
     * @param mixed $id
     * @param string $name
     */
    protected function generateTABLECOLUMN($id = '2', $name = '')
    {
        $xml = '<tableColumn id="' . $id . '" name="' . $name . '"></tableColumn >__PHX=__GENERATETABLECOLUMNS__';

        $this->xml = str_replace('__PHX=__GENERATETABLECOLUMNS__', $xml, $this->xml);
    }

    /**
     * Generate tablecolumns
     *
     * @access protected
     * @param mixed $count
     */
    protected function generateTABLECOLUMNS($count = '2')
    {
        $xml = '<tableColumns count="' . $count . '">__PHX=__GENERATETABLECOLUMNS__</tableColumns>__PHX=__GENERATETABLE__';

        $this->xml = str_replace('__PHX=__GENERATETABLE__', $xml, $this->xml);
    }

    /**
     * Generate tablestyleinfo
     *
     * @access protected
     * @param string $showFirstColumn
     * @param string $showLastColumn
     * @param string $showRowStripes
     * @param string $showColumnStripes
     */
    protected function generateTABLESTYLEINFO($showFirstColumn = '0', $showLastColumn = "0", $showRowStripes = "1", $showColumnStripes = "0")
    {
        $xml = '<tableStyleInfo showFirstColumn="' . $showFirstColumn . '" showLastColumn="' . $showLastColumn . '" showRowStripes="' . $showRowStripes . '" showColumnStripes="' . $showColumnStripes . '"></tableStyleInfo >';

        $this->xml = str_replace('__PHX=__GENERATETABLE__', $xml, $this->xml);
    }
}
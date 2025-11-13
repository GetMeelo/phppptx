<?php

/**
 * Create doughnut chart
 *
 * @category   Phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateChartDoughnut extends CreateChartElement
{
    /**
     * Create embedded xml chart
     *
     * @access public
     */
    public function createEmbeddedXmlChart()
    {
        $this->xmlChart = '';
        $this->generateCHARTSPACE();
        $this->generateDATE1904(1);
        $this->generateLANG();
        $this->generateROUNDEDCORNERS($this->roundedCorners);
        $color = 2;
        if (!empty($this->color)) {
            $color = $this->color;
        }
        $this->generateSTYLE($color);
        $this->generateCHART();
        if ($this->title != '') {
            $this->generateTITLE();
            $this->generateTITLETX();
            $this->generateRICH();
            $this->generateBODYPR();
            $this->generateLSTSTYLE();
            $this->generateTITLEP();
            $this->generateTITLEPPR();
            $this->generateDEFRPR('title');
            $this->generateTITLER();
            $this->generateTITLERPR();
            $this->generateTITLET($this->title);
            $this->cleanTemplateFonts();
        } else {
            $this->generateAUTOTITLEDELETED();
            $title = '';
        }

        if (strpos($this->type, '3D') !== false) {
            $this->generateVIEW3D();
            $rotX = 30;
            $rotY = 30;
            $perspective = 30;
            if ($this->rotX != '') {
                $rotX = $this->rotX;
            }
            if ($this->rotY != '') {
                $rotY = $this->rotY;
            }
            if ($this->perspective != '') {
                $perspective = $this->perspective;
            }
            $this->generateROTX($rotX);
            $this->generateROTY($rotY);
            $this->generateRANGAX($this->rAngAx);
            $this->generatePERSPECTIVE($perspective);
        }
        $this->generatePLOTAREA();
        $this->generateLAYOUT();

        $this->generateDOUGHNUTCHART();
        $this->generateVARYCOLORS();
        $legends = array($this->title);
        $numValues = count($this->values['data']);
        $letter = 'A';
        for ($i = 0; $i < count($legends); $i++) {
            $this->generateSER();
            $this->generateIDX($i);
            $this->generateORDER($i);
            if (function_exists('str_increment')) {
                $letter = str_increment($letter);
            } else {
                $letter++;
            }

            $this->generateTX();
            $this->generateSTRREF();
            $this->generateF('Sheet1!$' . $letter . '$1');
            $this->generateSTRCACHE();
            $this->generatePTCOUNT();
            $this->generatePT();
            $this->generateV($legends[$i]);
            if (!empty($this->explosion) && is_numeric($this->explosion)) {
                $this->generateEXPLOSION($this->explosion);
            }
            $this->cleanTemplate2();

            if (is_array($this->theme) && isset($this->theme['serRgbColors']) && isset($this->theme['serRgbColors'][$i])) {
                if ($this->theme['serRgbColors'][$i] != null) {
                    $this->generateSPPR_SER();
                    $this->generateSPPR_SOLIDFILL($this->theme['serRgbColors'][$i]);
                }
            }

            if (is_array($this->theme) && isset($this->theme['valueRgbColors']) && isset($this->theme['valueRgbColors'][$i]) && $this->theme['valueRgbColors'][$i] != null) {
                if ($this->theme['valueRgbColors'][$i] != null) {
                    $this->generateCDPT($this->theme['valueRgbColors'][$i]);
                }
            }

            $this->generateCAT();
            $this->generateSTRREF();
            $this->generateF('Sheet1!$A$2:$A$' . ($numValues + 1));
            $this->generateSTRCACHE();
            $this->generatePTCOUNT($numValues);

            $num = 0;
            foreach ($this->values['data'] as $value) {
                $this->generatePT($num);
                $this->generateV($value['name']);
                $num++;
            }
            $this->cleanTemplate2();
            $this->generateVAL();
            $this->generateNUMREF();
            $this->generateF('Sheet1!$' . $letter . '$2:$' . $letter . '$' . ($numValues + 1));
            $this->generateNUMCACHE();
            $this->generateFORMATCODE();
            $this->generatePTCOUNT($numValues);
            $num = 0;
            foreach ($this->values['data'] as $name => $value) {
                $this->generatePT($num);
                $this->generateV($value['values'][$i]);
                $num++;
            }
            $this->cleanTemplate3();
        }

        //Generate labels
        $this->generateSERDLBLS();

        if ($this->formatCode) {
            $this->generateNUMFMT($this->formatCode, 0);
        }

        $this->generateSHOWLEGENDKEY($this->showLegendKey);
        $this->generateSHOWVAL($this->showValue);
        $this->generateSHOWCATNAME($this->showCategory);
        $this->generateSHOWSERNAME($this->showSeries);
        $this->generateSHOWPERCENT($this->showPercent);
        $this->generateSHOWBUBBLESIZE($this->showBubbleSize);
        $this->generateFIRSTSLICEANG();
        if (!empty($this->holeSize) && is_numeric($this->holeSize)) {
            $this->generateHOLESIZE($this->holeSize);
        } else {
            $this->generateHOLESIZE();
        }

        $this->generateLEGEND();
        $this->generateLEGENDPOS($this->legendPos);
        $this->generateLEGENDOVERLAY($this->legendOverlay);
        $this->generatePLOTVISONLY();

        if ((!isset($this->border) || $this->border == 0 || !is_numeric($this->border))
        ) {
            $this->generateSPPR();
            $this->generateLN();
            $this->generateNOFILL();
        } else {
            $this->generateSPPR();
            $this->generateLN($this->border);
        }

        if ($this->font != '') {
            $this->generateTXPR();
            $this->generateLEGENDBODYPR();
            $this->generateLSTSTYLE();
            $this->generateAP();
            $this->generateAPPR();
            $this->generateDEFRPR();
            $this->generateRFONTS($this->font);
            $this->generateENDPARARPR();
        }

        $this->generateEXTERNALDATA();
        $this->cleanTemplateDocument();

        return $this->xmlChart;
    }

    /**
     * Return the tag
     *
     * @access public
     * @return array
     */
    public function dataTag()
    {
        return array('val');
    }

    /**
     * Return the type of the xlsx object
     *
     * @access public
     */
    public function getXlsxType()
    {
        return new CreateCompletedXlsx();
    }

    /**
     * Generate w:numFmt
     *
     * @access protected
     * @param string $formatCode
     * @param mixed $sourceLinked
     */
    protected function generateNUMFMT($formatCode = 'General', $sourceLinked = '1')
    {
        $this->xmlChart = str_replace('__PHX=__GENERATEDLBLS__', '<c:numFmt formatCode="' . $formatCode . '" sourceLinked="' . $sourceLinked . '"></c:numFmt>__PHX=__GENERATEDLBLS__', $this->xmlChart);
    }
}
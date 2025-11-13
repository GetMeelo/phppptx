<?php

/**
 * Create line chart
 *
 * @category   Phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateChartLine extends CreateChartElement
{
    /**
     * Create embedded xml chart
     *
     * @access public
     */
    public function createEmbeddedXmlChart()
    {
        $this->xmlChart = '';
        $args = func_get_args();
        $this->generateCHARTSPACE();
        $this->generateDATE1904(1);
        $this->generateLANG();
        $color = 2;
        if ($this->color) {
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
            $this->generateTITLELAYOUT();
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

        if (strpos($this->type, '3D') !== false) {
            $this->generateLINE3DCHART();
        } else {
            $this->generateLINECHART();
        }
        $groupBar = 'standard';
        if (!empty($this->groupBar) && in_array($this->groupBar, array('stacked', 'standard', 'percentStacked'))) {
            $groupBar = $this->groupBar;
        }
        $this->generateGROUPING($groupBar);
        $legends = array();
        if (isset($this->values['legend'])) {
            $legends = $this->values['legend'];
        }

        $numValues = count($this->values['data']);
        $this->generateVARYCOLORS($this->varyColors);
        $letter = 'A';
        // keep the max id value to be used with combo charts
        $idxMax = 0;
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

            if (!empty($this->symbol)) {
                $sizeSimbol = NULL;
                if (is_numeric($this->symbolSize)) {
                    $sizeSimbol = $this->symbolSize;
                }
                if (is_array($this->symbol)) {
                    $this->generateMARKER($this->symbol[$i], $sizeSimbol);
                } else {
                    $this->generateMARKER($this->symbol, $sizeSimbol);
                }
            }
            $this->cleanTemplate2();

            if (isset($this->values['trendline'][$i])) {
                $this->generateTRENDLINE($this->values['trendline'][$i]);
            }

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

            if (is_array($this->theme) && isset($this->theme['serDataLabels'])) {
                if ($this->theme['serDataLabels'][$i] != null) {
                    $this->generateDATALABELS_SER($this->theme['serDataLabels'][$i], $i);
                }
            }

            if (is_array($this->theme) && isset($this->theme['valueDataLabels'])) {
                if (isset($this->theme['valueDataLabels'][$i]) && $this->theme['valueDataLabels'][$i] != null) {
                    $this->generateDATALABELS_DLBL($this->theme['valueDataLabels'][$i], $i);
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

            if (!empty($this->smooth)) {
                $this->generateSMOOTH();
            } else if ($this->smooth === '0') {
                $this->generateSMOOTH(0);
            }
            $this->cleanTemplate3();

            $idxMax = $i;
        }

        //Generate labels
        $this->generateSERDLBLS();
        $this->generateSHOWLEGENDKEY($this->showLegendKey);
        $this->generateSHOWVAL($this->showValue);
        $this->generateSHOWCATNAME($this->showCategory);
        $this->generateSHOWSERNAME($this->showSeries);
        $this->generateSHOWPERCENT($this->showPercent);
        $this->generateSHOWBUBBLESIZE($this->showBubbleSize);

        if ($this->groupBar != '' && ($this->groupBar == 'stacked' || $this->groupBar == 'percentStacked')) {
            $this->generateOVERLAP();
        }
        $this->generateAXID();
        $this->generateAXID(59040512);

        if (strpos($this->type, '3D') !== false) {
            $this->generateAXID(83319040);
        }
        $this->generateVALAX();
        $this->generateAXAXID(59040512);
        $this->generateSCALING(true);
        $this->generateDELETE($this->delete);
        if (!empty($this->orientation) && is_array($this->orientation) && isset($this->orientation[0]) && !is_null($this->orientation[0]))  {
            $this->generateORIENTATION($this->orientation[0]);
        } else {
            $this->generateORIENTATION();
        }
        if (!empty($this->axPos) && is_array($this->axPos) && isset($this->axPos[0]) && !is_null($this->axPos[0]))  {
            $this->generateAXPOS($this->axPos[0]);
        } else {
            $this->generateAXPOS('l');
        }

        switch ($this->hgrid) {
            case 1:
                $this->generateMAJORGRIDLINES();
                break;
            case 2:
                $this->generateMINORGRIDLINES();
                break;
            case 3:
                $this->generateMAJORGRIDLINES();
                $this->generateMINORGRIDLINES();
                break;
            default:
                break;
        }

        if (!empty($this->vaxLabel)) {
            $this->generateAXLABEL($this->vaxLabel);
            $vert = 'horz';
            $rot = 0;
            if ($this->vaxLabelDisplay == 'vertical') {
                $vert = 'wordArtVert';
            }
            if ($this->vaxLabelDisplay == 'rotated') {
                $rot = '-5400000';
            }
            $this->generateAXLABELDISP($vert, $rot);
        }
        if ($this->formatCode) {
            $this->generateNUMFMT($this->formatCode, 0);
        } else {
            $this->generateNUMFMT();
        }
        if (!is_array($this->tickLblPos)) {
            $this->generateTICKLBLPOS($this->tickLblPos, true);
        } else if (!empty($this->tickLblPos) && is_array($this->tickLblPos) && isset($this->tickLblPos[0]) && !is_null($this->tickLblPos[0])) {
            $this->generateTICKLBLPOS($this->tickLblPos[0]);
        }
        $this->generateMAJORUNIT($this->majorUnit);
        $this->generateMINORUNIT($this->minorUnit);
        $this->generateCROSSAX(59034624);
        $this->generateCROSSES();
        $this->generateCROSSBETWEEN();
        $this->generateCATAX();
        $this->generateAXAXID(59034624);
        $this->generateSCALING();
        $this->generateDELETE($this->delete);
        if (!empty($this->orientation) && is_array($this->orientation) && isset($this->orientation[1]) && !is_null($this->orientation[1]))  {
            $this->generateORIENTATION($this->orientation[1]);
        } else {
            $this->generateORIENTATION();
        }
        if (!empty($this->axPos) && is_array($this->axPos) && isset($this->axPos[1]) && !is_null($this->axPos[1]))  {
            $this->generateAXPOS($this->axPos[1]);
        } else {
            $this->generateAXPOS();
        }

        if (!empty($this->haxLabel)) {
            $this->generateAXLABEL($this->haxLabel);
            $vert = 'horz';
            $rot = 0;
            if ($this->haxLabelDisplay == 'vertical') {
                $vert = 'wordArtVert';
            }
            if ($this->haxLabelDisplay == 'rotated') {
                $rot = '-5400000';
            }
            $this->generateAXLABELDISP($vert, $rot);
        }
        switch ($this->vgrid) {
            case 1:
                $this->generateMAJORGRIDLINES();
                break;
            case 2:
                $this->generateMINORGRIDLINES();
                break;
            case 3:
                $this->generateMAJORGRIDLINES();
                $this->generateMINORGRIDLINES();
                break;
            default:
                break;
        }

        if (strpos($this->type, 'surface') !== false) {
            $this->generateMAJORTICKMARK();
        }
        if (!is_array($this->tickLblPos)) {
            $this->generateTICKLBLPOS();
        } else if (!empty($this->tickLblPos) && is_array($this->tickLblPos) && isset($this->tickLblPos[1]) && !is_null($this->tickLblPos[1])) {
            $this->generateTICKLBLPOS($this->tickLblPos[1]);
        }
        $this->generateCROSSAX();
        $this->generateCROSSES();
        $this->generateAUTO();
        $this->generateLBLALGN();
        $this->generateLBLOFFSET();
        if ($this->showTable) {
            $this->generateDATATABLE();
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
}
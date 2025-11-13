<?php

/**
 * Create scatter chart
 *
 * @category   Phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateChartScatter extends CreateChartElement
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
            $this->generatePERSPECTIVE($perspective);
        }
        $this->generatePLOTAREA();
        $this->generateLAYOUT();

        $this->generateSCATTERCHART();
        $this->generateSCATTERSTYLE($this->style);
        $this->generateVARYCOLORS($this->varyColors);
        // analyze data to check if it contains one or more series
        $seriesNumPt = array();
        // reorder data to keep values of each serie in the same array
        $seriesData = array();
        foreach ($this->values['data'] as $datas) {
            $i = 0;
            foreach ($datas as $dataValues) {
                if (!is_array($dataValues[0])) {
                    if (!isset($seriesNumPt[$i])) {
                        $seriesNumPt[$i] = 0;
                    }
                    $seriesNumPt[$i] += 1;
                    $seriesData[$i][] = $dataValues;
                } else if (is_array($dataValues[0])) {
                    $j = 0;
                    foreach ($dataValues as $dataValuesMultipleAxis) {
                        // avoid adding empty strings to allow omitting axis
                        if (count($dataValuesMultipleAxis) == 0) {
                            $j++;
                            continue;
                        }
                        if (!isset($seriesNumPt[$j])) {
                            $seriesNumPt[$j] = 0;
                        }
                        $seriesNumPt[$j] += 1;
                        $seriesData[$j][] = $dataValuesMultipleAxis;
                        $j++;
                    }
                }
            }
            $i++;
        }

        foreach ($seriesNumPt as $seriesNumPtIndex => $seriesNumPtCount) {
            $legend = 'Values';
            if (isset($this->values['legend']) && isset($this->values['legend'][$seriesNumPtIndex])) {
                $legend = $this->values['legend'][$seriesNumPtIndex];
            }
            $letter = 'A';
            $this->generateSER();
            $this->generateIDX($seriesNumPtIndex);
            $this->generateORDER($seriesNumPtIndex);
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
            $this->generateV($legend);
            if (!empty($this->symbol)) {
                if ($this->symbol == 'line') {
                    $this->generateMARKER('none');
                } elseif ($this->symbol == 'dot') {
                    $this->generateSPPR_SER();
                    $this->generateLN(2);
                    $this->generateNOFILL();
                }
            }

            if (is_array($this->theme) && isset($this->theme['serRgbColors']) && isset($this->theme['serRgbColors'][$seriesNumPtIndex])) {
                if ($this->theme['serRgbColors'][$seriesNumPtIndex] != null) {
                    $this->generateSPPR_SER();
                    $this->generateSPPR_SOLIDFILL($this->theme['serRgbColors'][$seriesNumPtIndex]);
                }
            }

            if (is_array($this->theme) && isset($this->theme['valueRgbColors']) && isset($this->theme['valueRgbColors'][$seriesNumPtIndex]) && $this->theme['valueRgbColors'][$seriesNumPtIndex] != null) {
                if ($this->theme['valueRgbColors'][$seriesNumPtIndex] != null) {
                    $this->generateCDPT($this->theme['valueRgbColors'][$seriesNumPtIndex]);
                }
            }

            if (is_array($this->theme) && isset($this->theme['serDataLabels'])) {
                if ($this->theme['serDataLabels'][$seriesNumPtIndex] != null) {
                    $this->generateDATALABELS_SER($this->theme['serDataLabels'][$seriesNumPtIndex], $seriesNumPtIndex);
                }
            }

            if (is_array($this->theme) && isset($this->theme['valueDataLabels'])) {
                if (isset($this->theme['valueDataLabels'][$seriesNumPtIndex]) && $this->theme['valueDataLabels'][$seriesNumPtIndex] != null) {
                    $this->generateDATALABELS_DLBL($this->theme['valueDataLabels'][$seriesNumPtIndex], $seriesNumPtIndex);
                }
            }

            $this->cleanTemplate2();

            $this->generateXVAL();
            $this->generateNUMREF();
            $this->generateF('Sheet1!$A$2:$A$' . ($seriesNumPtCount + 1));
            $this->generateNUMCACHE();
            $this->generateFORMATCODE();
            $this->generatePTCOUNT($seriesNumPtCount);
            $num = 0;
            foreach ($seriesData[$seriesNumPtIndex] as $data) {
                $this->generatePT($num);
                $this->generateV($data[0]);
                $num++;
            }
            $this->cleanTemplate2();
            $this->generateYVAL();
            $this->generateNUMREF();
            $this->generateF('Sheet1!$B$2:$B$' . ($seriesNumPtCount + 1));
            $this->generateNUMCACHE();
            $this->generateFORMATCODE();
            $this->generatePTCOUNT($seriesNumPtCount);
            $num = 0;
            foreach ($seriesData[$seriesNumPtIndex] as $data) {
                $this->generatePT($num);
                $this->generateV($data[1]);
                $num++;
            }
            $this->cleanTemplate2();
            if (!empty($this->smooth)) {
                $this->generateSMOOTH();
            } else if ($this->smooth === '0') {
                $this->generateSMOOTH(0);
            }
            $this->cleanTemplate3();
        }

        // generate labels
        $this->generateSERDLBLS();
        $this->generateSHOWLEGENDKEY($this->showLegendKey);
        $this->generateSHOWVAL($this->showValue);
        $this->generateSHOWCATNAME($this->showCategory);
        $this->generateSHOWSERNAME($this->showSeries);
        $this->generateSHOWPERCENT($this->showPercent);
        $this->generateSHOWBUBBLESIZE($this->showBubbleSize);

        $this->generateAXID();
        $this->generateAXID(59040512);
        $this->generateVALAX();
        $this->generateAXAXID(59034624);
        $this->generateSCALING();
        $this->generateDELETE($this->delete);
        if (!empty($this->orientation) && is_array($this->orientation) && isset($this->orientation[0]) && !is_null($this->orientation[0]))  {
            $this->generateORIENTATION($this->orientation[0]);
        } else {
            $this->generateORIENTATION();
        }
        if (!empty($this->axPos) && is_array($this->axPos) && isset($this->axPos[0]) && !is_null($this->axPos[0]))  {
            $this->generateAXPOS($this->axPos[0]);
        } else {
            $this->generateAXPOS();
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
        if ($this->formatCode) {
            $this->generateNUMFMT($this->formatCode, 0);
        } else {
            $this->generateNUMFMT();
        }
        if (!is_array($this->tickLblPos)) {
            $this->generateTICKLBLPOS();
        } else if (!empty($this->tickLblPos) && is_array($this->tickLblPos) && isset($this->tickLblPos[0]) && !is_null($this->tickLblPos[0])) {
            $this->generateTICKLBLPOS($this->tickLblPos[0]);
        }
        $this->generateCROSSAX();
        $this->generateCROSSES();
        $this->generateCROSSBETWEEN('midCat');
        $this->generateVALAX();
        $this->generateAXAXID(59040512);
        $this->generateSCALING(true);
        $this->generateDELETE($this->delete);
        if (!empty($this->orientation) && is_array($this->orientation) && isset($this->orientation[1]) && !is_null($this->orientation[1]))  {
            $this->generateORIENTATION($this->orientation[1]);
        } else {
            $this->generateORIENTATION();
        }
        if (!empty($this->axPos) && is_array($this->axPos) && isset($this->axPos[1]) && !is_null($this->axPos[1]))  {
            $this->generateAXPOS($this->axPos[1]);
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
        $this->generateNUMFMT();
        if (!is_array($this->tickLblPos)) {
            $this->generateTICKLBLPOS($this->tickLblPos, true);
        } else if (!empty($this->tickLblPos) && is_array($this->tickLblPos) && isset($this->tickLblPos[1]) && !is_null($this->tickLblPos[1])) {
            $this->generateTICKLBLPOS($this->tickLblPos[1]);
        }
        $this->generateMAJORUNIT($this->majorUnit);
        $this->generateMINORUNIT($this->minorUnit);
        $this->generateCROSSAX(59034624);
        $this->generateCROSSES();
        $this->generateCROSSBETWEEN();

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
        return array('xVal', 'yVal');
    }

    /**
     * Return the type of the xlsx object
     *
     * @access public
     */
    public function getXlsxType()
    {
        return new CreateScatterXlsx();
    }
}
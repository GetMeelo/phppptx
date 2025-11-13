<?php

/**
 * Create math eq
 *
 * @category   Phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateMath extends CreateElement
{
    /**
     * Generate a new MathML equation
     *
     * @access public
     * @param string $equation OMML equation string or MathML
     * @param string $type Type of equation: omml, mathml
     * @param array $options
     *      'align' (string) left, center, right
     *      'bold' (bool)
     *      'color' (string) ffffff, ff0000...
     *      'fontSize' (int) 8, 9, 10...
     *      'isPptxFragment' (bool)
     *      'italic' (bool)
     *      'underline' (string) : single...
     *  @return string
     */
    public function createElementMath($equation, $type, $options = array())
    {
        $stylesMathEq = '';
        if (isset($options['align'])) {
            $stylesMathEq = '<m:oMathParaPr xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:jc m:val="'.$options['align'].'"/></m:oMathParaPr>';
        }

        $newEquation = '';

        if ($type == 'omml') {
            // apply styles if exist
            if (is_array($options) && count($options) > 0) {
                $equation = $this->mathEquationStyles($equation, $options);
            }

            $newEquation = '<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' . $equation . '</a14:m>';
            if (!$options['isPptxFragment']) {
                $newEquation = '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' . $newEquation . '</a:p>';
            }
        } elseif ($type == 'mathml') {
            // transform the MathML equation to OOML
            $equation = $this->transformMath($equation);

            // apply styles if exist
            if (is_array($options) && count($options) > 0) {
                $equation = $this->mathEquationStyles($equation, $options);
            }

            $newEquation = '<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' . $stylesMathEq . $equation . '</m:oMathPara></a14:m>';
            if (!$options['isPptxFragment']) {
                $newEquation = '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' . $newEquation . '</a:p>';
            }
        }

        return $newEquation;
    }

    /**
     * Apply math equation styles
     *
     * @param string $equationStyles
     * @param array $options
     * @return string
     */
    protected function mathEquationStyles($equationStyles, $options)
    {
        $xmlUtilities = new XmlUtilities();
        $equationStyledDOM = $xmlUtilities->generateDomDocument($equationStyles);

        $elementsMR = $equationStyledDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/officeDocument/2006/math', 'r');
        if ($elementsMR->length > 0) {
            foreach ($elementsMR as $elementMR) {
                $elementRPR = $elementMR->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'rPr');
                // a:rPR doesn't exist, create it
                if ($elementRPR->length == 0) {
                    $elementTag = $elementMR->ownerDocument->createElement('a:rPr');
                    $elementMR->insertBefore($elementTag, $elementMR->firstChild);

                    $elementRPRItem = $elementTag;
                } else {
                    $elementRPRItem = $elementRPR->item(0);
                }

                if (isset($options['color'])) {
                    $elementSolidFill = $elementMR->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'solidFill');
                    // a:solidFill doesn't exist, create and add it
                    if ($elementSolidFill->length == 0) {
                        $elementTagSolidFill = $elementRPRItem->ownerDocument->createElement('a:solidFill');
                        $elementTagSrgbClr = $elementRPRItem->ownerDocument->createElement('a:srgbClr');
                        $elementTagSrgbClr->setAttribute('val', $options['color']);
                        $elementTagSolidFill->appendChild($elementTagSrgbClr);
                        $elementRPRItem->appendChild($elementTagSolidFill);
                    }
                }

                if (isset($options['fontSize'])) {
                    $elementRPRItem->setAttribute('sz', (string)((int)$options['fontSize'] * 100));
                }
            }
        }

        return $equationStyledDOM->saveXML($equationStyledDOM->documentElement);
    }

    /**
     * Transform a MathML eq using XSL
     *
     * @access protected
     * @param string $mathML MathML equation
     * @return string OMML equation
     */
    protected function transformMath($mathML)
    {
        $xmlUtilities = new XmlUtilities();
        $rscXML = $xmlUtilities->generateDomDocument($mathML);

        $objXSLTProc = new XSLTProcessor();
        $objXSL = new DOMDocument();
        $objXSL->load(__DIR__ . '/../xsl/MML2OMML_n.XSL');
        $objXSLTProc->importStylesheet($objXSL);

        $newEquation = $objXSLTProc->transformToXML($rscXML);
        $arrOMML = array('<?xml version="1.0" encoding="UTF-8"?>');
        $newEquation = str_replace($arrOMML, '', $newEquation);

        return $newEquation;
    }
}
<?php

/**
 * Create table style
 *
 * @category   Phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateTableStyle
{
    /**
     * Construct
     *
     * @access public
     */
    public function __construct() {}

    /**
     * Creates table style
     *
     * @access public
     * @param string $name
     * @param array $styles
     * @return string
     */
    public function createTableStyle($name, $styles)
    {
        $guid = PhppptxUtilities::generateGUID();

        // keep correct orders
        $validBorders = array('left', 'right', 'top', 'bottom', 'insideH', 'insideV');
        $tableTypes = array('wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'lastCol', 'firstCol', 'lastRow', 'seCell', 'swCell', 'firstRow', 'neCell', 'nwCell');

        $tableStyle = '<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="'.$guid['guid'].'" styleName="'.$name.'">';

        foreach ($tableTypes as $tableType) {
            $styleValue = array();
            if (isset($styles[$tableType])) {
                $styleValue = $styles[$tableType];
            }

            $tableStyle .= '<a:'.$tableType.'>';

            $tcTxInlineStyles = '';
            if (isset($styleValue['bold']) || isset($styleValue['italic'])) {
                if (isset($styleValue['bold']) && $styleValue['bold']) {
                    $tcTxInlineStyles .= ' b="on" ';
                }
                if (isset($styleValue['italic']) && $styleValue['italic']) {
                    $tcTxInlineStyles .= ' i="on" ';
                }
            }
            $tcTxBlockStyles = '';

            $tableStyle .= '<a:tcTxStyle '.$tcTxInlineStyles.'>'.$tcTxBlockStyles.'</a:tcTxStyle>';
            $tableStyle .= '<a:tcStyle>';
            if (isset($styleValue['border'])) {
                $tableStyle .= '<a:tcBdr>';
                foreach ($validBorders as $borderKey) {
                    if (isset($styleValue['border'][$borderKey])) {
                        $tableStyle .= '<a:'.$borderKey.'>';
                        if (!isset($styleValue['border'][$borderKey]['width'])) {
                            $styleValue['border'][$borderKey]['width'] = 12700;
                        }
                        $tableStyle .= '<a:ln w="'.$styleValue['border'][$borderKey]['width'].'">';
                        if (!isset($styleValue['border'][$borderKey]['color'])) {
                            $styleValue['border'][$borderKey]['color'] = '000000';
                        }
                        if ($styleValue['border'][$borderKey]['color'] == 'none') {
                            $tableStyle .= '<a:noFill/>';
                        } else {
                            $tableStyle .= '<a:solidFill><a:srgbClr val="'.str_replace('#', '', $styleValue['border'][$borderKey]['color']).'"/></a:solidFill>';
                        }
                        if (isset($styleValue['border'][$borderKey]['dash'])) {
                            $tableStyle .= '<a:prstDash val="'.$styleValue['border'][$borderKey]['dash'].'"/>';
                        }
                        $tableStyle .= '</a:ln>';

                        $tableStyle .= '</a:'.$borderKey.'>';
                    }
                }
                $tableStyle .= '</a:tcBdr>';
            } else {
                $tableStyle .= '<a:tcBdr/>';
            }
            if (isset($styleValue['backgroundColor'])) {
                $tableStyle .= '<a:fill>';
                $tableStyle .= '<a:solidFill><a:srgbClr val="'.str_replace('#', '', $styleValue['backgroundColor']).'"/></a:solidFill>';
                $tableStyle .= '</a:fill>';
            }
            $tableStyle .= '</a:tcStyle>';
            $tableStyle .= '</a:'.$tableType.'>';
        }

        $tableStyle .= '</a:tblStyle>';

        return $tableStyle;
    }
}
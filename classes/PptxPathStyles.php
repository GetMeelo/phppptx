<?php

/**
 * PptxPath extra functions
 *
 * @category   Phppptx
 * @package    PptxPath
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class PptxPathStyles
{
    /**
     * Analyzes variables in contents
     *
     * @access public
     * @param string $variable
     * @param string $xml
     * @return array
     */
    public function analyzeVariable($variable, $xml)
    {
        $variableType = array();

        $xmlUtilities = new XmlUtilities();
        $domDocument = $xmlUtilities->generateDomDocument($xml);
        $pptXPath = new DOMXPath($domDocument);
        $pptXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $pptXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

        // check if text content
        $domTexts = $pptXPath->query('//a:p[contains(.,"' . $variable . '")]');
        if ($domTexts->length > 0) {
            $variableType[] = 'text';
        }

        // check if table content
        $domTables = $pptXPath->query('//a:tc[.//a:p[contains(.,"' . $variable . '")]]');
        if ($domTables->length > 0) {
            $variableType[] = 'table';
        }

        // check if image content
        $domImages = $pptXPath->query('//p:pic[.//p:cNvPr[@descr="'.$variable.'" or @title="'.$variable.'"] and .//a:blip and not(.//a:videoFile) and not(.//a:audioFile)]');
        if ($domImages->length > 0) {
            $variableType[] = 'image';
        }

        // check if audio content
        $domAudios = $pptXPath->query('//p:pic[.//p:cNvPr[@descr="'.$variable.'" or @title="'.$variable.'"] and .//a:audioFile]');
        if ($domAudios->length > 0) {
            $variableType[] = 'audio';
        }

        // check if video content
        $domVideos = $pptXPath->query('//p:pic[.//p:cNvPr[@descr="'.$variable.'" or @title="'.$variable.'"] and .//a:videoFile]');
        if ($domVideos->length > 0) {
            $variableType[] = 'video';
        }

        return $variableType;
    }

    /**
     * Creates the required XML parser
     *
     * @access public
     * @param DOMNode $node
     * @return array
     */
    public function xmlParserStyle($node)
    {
        if ($node) {
            $parserXML = xml_parser_create();
            xml_parser_set_option($parserXML, XML_OPTION_CASE_FOLDING, 0);
            xml_parse_into_struct($parserXML, $node->ownerDocument->saveXML($node), $values, $indexes);

            return $values;
        } else {
            return array();
        }
    }
}
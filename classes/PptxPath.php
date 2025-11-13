<?php

/**
 * Generate XPath queries to select content in a PPTX
 *
 * @category   Phppptx
 * @package    PptxPath
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class PptxPath
{
    /**
     * Creates the required XPath query expression
     *
     * @access public
     * @param string $type audio, chart, diagram, image, paragraph, run, section, shape (text box), slide, table, table-row, table-cell, table-cell-paragraph, video
     * @param array $filters
     *      'contains' (string)
     *      'occurrence' (int) exact occurrence, (array) occurrences, (string) range of contents (e.g.: 2..9, 2.., ..9) or first() or last()
     *      'attributes' (array) node or descendant attributes
     *      'parent' (string) immediate children (default as '/', any parent) or any other parent (a:tbl/, p:sp/...)
     *      'rootParent' (string) root parent. Default as p:spTree for slide elements and p:presentation for presentation elements
     *      'target' (string) slides (default)
     * @param array $options
     * @return string
     */
    public static function xpathContentQuery($type, $filters, $options = array())
    {
        $contentTypes = array(
            'audio' => 'p:pic[.//a:audioFile]',
            'chart' => 'p:graphicFrame[.//c:chart]',
            'diagram' => 'p:graphicFrame[.//a:graphicData[@uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"]]',
            'image' => 'p:pic[.//a:blip and not(.//a:videoFile) and not(.//a:audioFile)]',
            'paragraph' => 'p:txBody/a:p',
            'run' => 'a:r',
            'section' => 'p14:sectionLst/p14:section',
            'shape' => 'p:sp',
            'slide' => 'p:sldIdLst/p:sldId',
            'table' => 'p:graphicFrame[./a:graphic//a:tbl]',
            'table-row' => 'a:tr',
            'table-cell' => 'a:tc',
            'table-cell-paragraph' => 'a:tc//a:p',
            'video' => 'p:pic[.//a:videoFile]',
        );

        $nodeType = $contentTypes[$type];

        // set root parent
        $rootParent = 'p:spTree';
        switch ($type) {
            case 'audio':
            case 'chart':
            case 'image':
            case 'math':
            case 'paragraph':
            case 'run':
                $rootParent = 'p:spTree';
                break;
            case 'section':
            case 'slide':
                $rootParent = 'p:presentation';
                break;
            default:
                $rootParent = 'p:spTree';
                break;
        }

        // default as "/"
        if (!isset($filters['parent'])) {
            $filters['parent'] = '/';
        }

        $nodeType = $filters['parent'] . $nodeType;

        if (isset($filters['rootParent'])) {
            $rootParent = $filters['rootParent'];
        }

        $condition = '1=1';

        $lastCondition = '';

        if (isset($filters['contains'])) {
            $contentFilter =  ' contains(., \'' . $filters['contains']  . '\')';

            $condition .= ' and ' . $contentFilter;
        }

        if (isset($filters['attributes']) && is_array($filters['attributes'])) {
            // the attribute value may be a string if getting the current element
            // or an array if getting a descendant of the current element
            foreach ($filters['attributes'] as $keyAttribute => $valueAttribute) {
                if (is_array($valueAttribute)) {
                    // get the descendant of the current element based on the key attribute to get the descendant
                    foreach ($valueAttribute as $keyValue => $valueValue) {
                        $condition .= ' and descendant::'.$keyAttribute.'[contains(@'.$keyValue.', "'.$valueValue.'")]';
                    }
                } else {
                    // get the current element
                    $condition .= ' and contains(@'.$keyAttribute.', "'.$valueAttribute.'")';
                }
            }
        }

        $mainQuery = '//' . $rootParent . '/' . $nodeType . '[' . $condition . ']' . $lastCondition;

        // occurrence
        // if first() set value as 1
        if (isset($filters['occurrence']) && $filters['occurrence'] === 'first()') {
            $filters['occurrence'] = 1;
        }
        // if last() set value as -1
        if (isset($filters['occurrence']) && $filters['occurrence'] === 'last()') {
            $filters['occurrence'] = -1;
        }

        if (isset($filters['occurrence']) && is_int($filters['occurrence'])) {
            // position element
            $occurrence = ($filters['occurrence'] < 0) ? 'last()' : $filters['occurrence'];
            $mainQuery = '(' . $mainQuery . ')[' . $occurrence . ']';
        } else if (isset($filters['occurrence']) && is_array($filters['occurrence'])) {
            $rangeQuery = '[' . implode(' or ', array_map(function ($value) {
                return 'position() = ' . $value;
            }, $filters['occurrence'])) . ']';

            $mainQuery = '(' . $mainQuery . ')' . $rangeQuery;

        } elseif (isset($filters['occurrence'])) {
            // range elements
            $rangeValues = explode('..', $filters['occurrence']);

            // create the range query dynamically
            $rangeQuery = '[';

            // from
            if (isset($rangeValues[0]) && !empty($rangeValues[0])) {
                $rangeQuery .= 'position() >= ' . $rangeValues[0];
            }

            // to
            if (isset($rangeValues[1]) && !empty($rangeValues[1])) {
                if (isset($rangeValues[0]) && !empty($rangeValues[0])) {
                    $rangeQuery .= ' and ';
                }
                $rangeQuery .= 'position() <= ' . $rangeValues[1];
            }

            $rangeQuery .= ']';

            $mainQuery = '(' . $mainQuery . ')' . $rangeQuery;
        }

        $query = $mainQuery;

        return $query;
    }
}
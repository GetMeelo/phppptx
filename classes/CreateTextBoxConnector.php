<?php

/**
 * Create text box connector
 *
 * @category   Phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateTextBoxConnector extends CreateElement
{
    /**
     * Generate a new text box connector
     *
     * @access public
     * @param DOMDocument $slideDOM
     * @param array $position
     *      'coordinateX' (int) EMUs (English Metric Unit)
     *      'coordinateY' (int) EMUs (English Metric Unit)
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     *      'name' (string) internal name. If not set, a random name is generated
     *      'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $connection
     *      'start' (int|string) id or internal name
     *      'end' (int|string) id or internal name
     *      'positionStart' (string) connection position: top, left, right (default), bottom. Only used when $position is calculated automatically. If not set, automatically detect the best position
     *      'positionEnd' (string) connection position: top, left (default), right, bottom. Only used when $position is calculated automatically. If not set, automatically detect the best position
     *      'flipH' (bool) flipped horizontally. Default as false
     *      'flipV' (bool) flipped vertically. Default as false
     * @param array $options
     *      'color' (string) FF0000, 00FFFF,...
     *      'dash' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *      'geom' (string) bentConnector2, bentConnector3, bentConnector4, bentConnector5, curvedConnector2, curvedConnector3, curvedConnector4, curvedConnector5, straightConnector1 (default)
     *      'lineWidth' (int) EMUs (English Metric Unit). 12700 = 1pt
     *      'rId' (string) shape ID
     *      'rotation' (int) 60.000ths of a degree
     *      'shapeGuide' (array)
     *          'fmla' (string) shape guide formula
     *          'guide' (string) shape guide name
     *      'tailEnd' (string) arrow, diamond, none, oval, stealth, triangle (default)
     * @throws Exception position not valid
     * @return DOMNode
     */
    public function addElementTextBoxConnector($slideDOM, $position, $connection, $options = array())
    {
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

        if (!isset($position['coordinateX']) || !isset($position['coordinateY']) || !isset($position['sizeX']) || !isset($position['sizeY'])) {
            PhppptxLogger::logger('The chosen position is not valid. Use a valid position.', 'fatal');
        }

        $name = 'Shape Connector ' . $options['rId'];
        if (isset($position['name'])) {
            $name = $this->parseAndCleanTextString($position['name']);
        }

        // shape id. Generate a new random one that is not duplicated in the current slide
        $cNvPrId = null;
        while (!isset($cNvPrId)) {
            $randomId = rand(999999, 999999999);
            $nodesCnvPrId = $slideXPath->query('//p:cNvPr[@id="'.$randomId.'"]');
            if ($nodesCnvPrId->length == 0) {
                $cNvPrId = $randomId;
            }
        }

        // position in the text box
        // 0: top, 1: left, 2: bottom, 3: right
        $idxStart = 3;
        switch ($connection['positionStart']) {
            case 'top':
                $idxStart = 0;
                break;
            case 'left':
                $idxStart = 1;
                break;
            case 'bottom':
                $idxStart = 2;
                break;
            case 'right':
                $idxStart = 3;
                break;
            default:
                $idxStart = 3;
                break;
        }
        $idxEnd = 1;
        switch ($connection['positionEnd']) {
            case 'top':
                $idxEnd = 0;
                break;
            case 'left':
                $idxEnd = 1;
                break;
            case 'bottom':
                $idxEnd = 2;
                break;
            case 'right':
                $idxEnd = 3;
                break;
            default:
                $idxEnd = 1;
                break;
        }

        $flipped = '';
        if (isset($connection['flipV']) && $connection['flipV']) {
            $flipped = ' flipV="1"';
        }
        if (isset($connection['flipH']) && $connection['flipH']) {
            $flipped .= ' flipH="1"';
        }
        $rotation = '';
        if (isset($options['rotation'])) {
            $rotation = ' rot="'.$options['rotation'].'"';
        }

        $prstGeomContents = '<a:prstGeom prst="'.$options['geom'].'"><a:avLst/></a:prstGeom>';
        if (isset($options['shapeGuide']) && is_array($options['shapeGuide']) && count($options['shapeGuide']) > 0) {
            $prstGeomContents = '<a:prstGeom prst="'.$options['geom'].'"><a:avLst>';
            foreach ($options['shapeGuide'] as $shapeGuide) {
                if (isset($shapeGuide['fmla']) && isset($shapeGuide['guide'])) {
                    // add the shape guide
                    $prstGeomContents .= '<a:gd name="'.$shapeGuide['guide'].'" fmla="'.$shapeGuide['fmla'].'"/>';
                }
            }
            $prstGeomContents .= '</a:avLst></a:prstGeom>';
        }

        $styles = '<a:ln>';
        if (isset($options['lineWidth'])) {
            $styles = '<a:ln w="'.$options['lineWidth'].'">';
        }
        if (isset($options['color'])) {
            $styles .= '<a:solidFill><a:srgbClr val="'.$options['color'].'"/></a:solidFill>';
        }
        if (isset($options['dash'])) {
            $styles .= '<a:prstDash val="'.$options['dash'].'"/>';
        }
        $styles .= '<a:tailEnd type="'.$options['tailEnd'].'"/></a:ln>';

        $connectorShape = '<p:spPr><a:xfrm'.$flipped.$rotation.'><a:off x="'.$position['coordinateX'].'" y="'.$position['coordinateY'].'"/><a:ext cx="'.$position['sizeX'].'" cy="'.$position['sizeY'].'"/></a:xfrm>'.$prstGeomContents.$styles.'</p:spPr>';

        $textBoxConnectonrContent = '<p:cxnSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvCxnSpPr><p:cNvPr id="'.$cNvPrId.'" name="'.$name.'"></p:cNvPr><p:cNvCxnSpPr><a:cxnSpLocks/><a:stCxn id="'.$connection['start'].'" idx="'.$idxStart.'"/><a:endCxn id="'.$connection['end'].'" idx="'.$idxEnd.'"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>'.$connectorShape.'<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:cxnSp>';

        // insert the new content
        $nodeShape = $this->insertNewContentOrder($textBoxConnectonrContent, $position, $slideDOM, $slideXPath);

        return $nodeShape;
    }
}
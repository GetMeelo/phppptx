<?php

/**
 * Create a PPTX file
 *
 * @category   Phppptx
 * @package    create
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
require_once __DIR__ . '/AutoLoader.php';
AutoLoader::load();
PhppptxLogger::initErrorLevel();
require_once __DIR__ . '/Phppptx_config.php';

class CreatePptx
{
    const PHPPPTX_VERSION = '4.0';

    /**
     *
     * @access public
     * @static
     * @var bool
     */
    public static $cleanUTF8 = true;

    /**
     *
     * @access public
     * @static
     * @var array
     */
    public static $elementsId = array();

    /**
     *
     * @access public
     * @var bool
     * @static
     */
    public static $returnPptxStructure = false;

    /**
     *
     * @access public
     * @var bool
     * @static
     */
    public static $rtl;

    /**
     *
     * @access public
     * @var bool
     * @static
     */
    public static $streamMode = false;

    /**
     * @access public
     * @var array
     */
    public $activeSlide;

    /**
     *
     * @access protected
     * @var boolean
     */
    protected $isMacro;

    /**
     *
     * @access protected
     * @var DOMDocument
     */
    protected $pptxContentTypesDOM;

    /**
     *
     * @access protected
     * @var DOMDocument
     */
    protected $pptxRelsPresentationDOM;

    /**
     *
     * @access protected
     * @var DOMDocument
     */
    protected $pptxPresentationDOM;

    /**
     *
     * @access protected
     * @var array
     */
    protected $namespaces;

    /**
     *
     * @access protected
     * @var array
     */
    protected $phppptxconfig;

    /**
     * Paths of temps files to use as DOCX file
     *
     * @access protected
     * @var array
     */
    protected $tempFileXLSX;

    /**
     * XmlUtilities
     *
     * @access protected
     * @var XmlUtilities XML Utilities classes
     */
    protected $xmlUtilities;

    /**
     *
     * @access protected
     * @var PptxStructure
     */
    protected $zipPptx;

    /**
     *
     * @access public
     * @var string
     */
    public $target = 'document';

    /**
     * Constructor
     *
     * @access public
     * @param array $options
     *      'layout' (string) layout to be used in the first slide when the PPTX is created from scratch. Default as 'Title Slide'
     *      'section' (string) add a new section and assign the default first slide. Set the section name
     * @param string|PptxStructure $pptxTemplatePath user custom template (preserves PPTX content)
     * @throws Exception empty or not valid template
     * @throws Exception layout name doesn't exist
     */
    public function __construct($options = array(), $pptxTemplatePath = null)
    {
        // default options
        if (!isset($options['layout'])) {
            $options['layout'] = 'Title Slide';
        }

        // general settings
        $this->phppptxconfig = PhppptxUtilities::parseConfig();

        if (empty($pptxTemplatePath)) {
            // default base template
            $templateStructure = new PptxStructureTemplate();
            $this->zipPptx = $templateStructure->getStructure($options);

            PhppptxLogger::logger('Default base template.', 'info');
        } elseif ($pptxTemplatePath instanceof PptxStructure) {
            // PptxStructure object
            $this->zipPptx = $pptxTemplatePath;

            PhppptxLogger::logger('PptxStructure template.', 'info');
        } else {
            // template
            $this->zipPptx = new PptxStructure();
            $this->zipPptx->parsePptx($pptxTemplatePath);

            PhppptxLogger::logger('Custom template.', 'info');
        }
        // initialize some required variables
        $this->xmlUtilities = new XmlUtilities();
        $this->tempFileXLSX = array();
        $this->isMacro = false;
        self::$elementsId = array();

        // set as active slide the first one
        $this->activeSlide = array();
        $this->activeSlide['position'] = 0;
        // set as active slide the last one
        //$slidesContents = $this->zipPptx->getSlides();
        //$this->activeSlide['position'] = count($slidesContents) - 1;

        $this->pptxContentTypesDOM = $this->zipPptx->getContent('[Content_Types].xml', 'DOMDocument');
        $this->pptxPresentationDOM = $this->zipPptx->getContent('ppt/presentation.xml', 'DOMDocument');

        // get current namespaces
        $this->namespaces = $this->zipPptx->getNamespaces();

        // include the standard image defaults
        $this->generateDefault('gif', 'image/gif');
        $this->generateDefault('jpg', 'image/jpg');
        $this->generateDefault('png', 'image/png');
        $this->generateDefault('jpeg', 'image/jpeg');
        $this->generateDefault('bmp', 'image/bmp');

        // get the rels file
        $this->pptxRelsPresentationDOM = $this->zipPptx->getContent('ppt/_rels/presentation.xml.rels', 'DOMDocument');

        // set rtl static variable
        self::$rtl = false;
        if (isset($this->phppptxconfig['settings']['rtl'])) {
            PhppptxLogger::logger('RTL mode enabled in settings.', 'info');
            if ($this->phppptxconfig['settings']['rtl'] == 'true' || $this->phppptxconfig['settings']['rtl'] == '1') {
                self::$rtl = true;
            }
        }

        // zip stream mode
        if (isset($this->phppptxconfig['settings']['stream'])) {
            PhppptxLogger::logger('Stream mode enabled in settings.', 'info');
            if (($this->phppptxconfig['settings']['stream'] == 'true' || $this->phppptxconfig['settings']['stream'] == '1') && file_exists(__DIR__ . '/ZipStream.php')) {
                self::$streamMode = true;
            }
        }
    }

    /**
     * Gets the active slide
     *
     * @access public
     * @return array
     */
    public function getActiveSlide()
    {
        return $this->activeSlide;
    }

    /**
     * Adds an audio
     *
     * @param string $audio audio path. Audio formats: flac, mp3, wav, wma
     * @param array $position
     *      'new' (array) a new position is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $audioStyles
     *      'image' (array)
     *          'image' (string) image to be used as preview. Set a default one if not set
     *          'mime' (string) forces a mime (image/jpg, image/jpeg, image/png, image/gif, image/webp)
     *      'mime' (string) forces a mime (audio/mpeg, audio/x-wav, audio/x-ms-wma, audio/unknown)
     * @param array $options
     *      'descr' (string) alt text (descr) value
     * @throws Exception audio doesn't exist
     * @throws Exception audio format is not supported
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     * @throws Exception position not valid
     */
    public function addAudio($audio, $position, $audioStyles = array(), $options = array())
    {
        // get audio information
        $audioInformation = new AudioUtilities();
        $audioContents = $audioInformation->returnAudioContents($audio, $options);
        $audioStyles['contents'] = $audioContents;
        $options['type'] = 'audio';

        // add the audio
        $this->addMedia($audioContents, $position, $audioStyles, $options);

        PhppptxLogger::logger('Add audio.', 'info');
    }

    /**
     * Adds a background image
     *
     * @access public
     * @param mixed $image Image path, base64, stream or GdImage. Image formats: png, jpg, jpeg, gif
     * @param array $options
     *      'overwrite' (bool) if true overwrites the existing background image if it exists. Default as true
     *      'tilePictureAsTexture' (bool) default as false
     *      'transparency' (int) from 0 to 100
     * @throws Exception image doesn't exist
     */
    public function addBackgroundImage($image, $options = array())
    {
        // get image information
        $imageInformation = new ImageUtilities();
        $imageContents = $imageInformation->returnImageContents($image, $options);

        // default options
        if (!isset($options['overwrite'])) {
            $options['overwrite'] = true;
        }
        if (!isset($options['tilePictureAsTexture'])) {
            $options['tilePictureAsTexture'] = false;
        }

        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);

        // handle if the image is going to be added
        $addBackgroundImage = true;

        // check if the slide contains some image and it must be overwritten
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $nodesBgToRemove = $slideXPath->query('//p:cSld/p:bg');
        if ($nodesBgToRemove->length > 0) {
            if (!$options['overwrite']) {
                $addBackgroundImage = false;
            }
        }

        // only add the image if requested
        if ($addBackgroundImage) {
            // remove p:cSld/p:bg tag if any exists
            if ($nodesBgToRemove->length > 0) {
                foreach ($nodesBgToRemove as $nodeBgToRemove) {
                    $nodeBgToRemove->parentNode->removeChild($nodeBgToRemove);
                }
            }

            // generate an identifier
            $backgroundId = $this->generateUniqueId();

            // remove p:cSld/p:bg tag if any exists
            $slideXPath = new DOMXPath($slideDOM);
            $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
            $nodesBgToRemove = $slideXPath->query('//p:cSld/p:bg');
            if ($nodesBgToRemove->length > 0) {
                foreach ($nodesBgToRemove as $nodeBgToRemove) {
                    $nodeBgToRemove->parentNode->removeChild($nodeBgToRemove);
                }
            }

            // make sure that there exists the corresponding content type
            $this->generateDefault($imageContents['extension'], $imageContents['mime']);
            // copy the image in the media folder
            $this->zipPptx->addContent('ppt/media/img' . $backgroundId . '.' . $imageContents['extension'], $imageContents['content']);
            // generate the relationship
            $newRelationship = '<Relationship Id="rId' . $backgroundId . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img' . $backgroundId . '.' . $imageContents['extension'] . '" />';

            // fill type
            if ($options['tilePictureAsTexture']) {
                // as texture
                $imageContentFill = '<a:tile algn="tl" flip="none" sx="100000" sy="100000" tx="0" ty="0"/>';
            } else {
                // default
                $imageContentFill = '<a:stretch><a:fillRect/></a:stretch>';
            }

            // transparency
            $imageContentTransparency = '<a:lum/>';
            if (isset($options['transparency'])) {
                $imageContentValue = 100000 - ((int)$options['transparency'] * 1000);
                $imageContentTransparency = '<a:alphaModFix amt="'.$imageContentValue.'"/><a:lum/>';
            }

            // add the new p:bg tag before p:spTree
            $newbgXML = '<p:bg xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId'.$backgroundId.'">'.$imageContentTransparency.'</a:blip><a:srcRect/>'.$imageContentFill.'</a:blipFill><a:effectLst/></p:bgPr></p:bg>';
            $newBgFragment = $slideDOM->createDocumentFragment();
            $newBgFragment->appendXML($newbgXML);

            $nodesSpTree = $slideXPath->query('//p:spTree');
            if ($nodesSpTree->length > 0) {
                $nodesSpTree->item(0)->parentNode->insertBefore($newBgFragment, $nodesSpTree->item(0));
            }

            // add the new relationsip
            $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
            $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
            $this->generateRelationship($slideRelsDOM, $newRelationship);

            // refresh contents
            $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());
            $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

            // free DOMDocument resources
            $slideRelsDOM = null;

            PhppptxLogger::logger('Add background image.', 'info');
        }

        // free DOMDocument resources
        $slideDOM = null;
    }

    /**
     * Adds a chart
     *
     * @param string $chart Chart type: area, area3D, bar, bar3D, bar3DCone, bar3DCylinder, bar3DPyramid, bubble, col, col3D, col3DCone, col3DCylinder, col3DPyramid, doughnut, line, line3D, ofPie, pie, pie3D, radar, scatter, surface
     * @param array $position
     *      'new' (array) a new shape is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $chartData Chart data
     * @param array $chartStyles
     *      'axPos' (array) position of the axis (r, l, t, b), each value of the array for each position (if a value if null avoids adding it)
     *      'color' (string) (1, 2, 3...) color scheme
     *      'font' (string) Arial, Times New Roman ...
     *      'formatCode' (string) number format
     *      'formatDataLabels' (array)
     *          'rotation' => (int)
     *          'position' => (string) center, insideEnd, insideBase, outsideEnd
     *      'haxLabel' (string) horizontal axis label
     *      'haxLabelDisplay' (string) rotated, vertical, horizontal
     *      'hgrid' (int) 0 (no grid) 1 (only major grid lines - default) 2 (only minor grid lines) 3 (both major and minor grid lines)
     *      'legendOverlay' (bool) if true the legend may overlay the chart
     *      'legendPos' (string) r, l, t, b, none
     *      'majorUnit' (float) bar, col, line charts
     *      'minorUnit' (float) bar, col, line charts
     *      'orientation' (array) orientation of the axis, from min to max (minMax) or max to min (maxMin), each value of the array for each axis (if a value if null avoids adding it)
     *      'scalingMax' (float) scaling max value bar, col, line charts
     *      'scalingMin' (float) scaling min value bar, col, line charts
     *      'showCategory' (bool) shows the categories inside the chart
     *      'showLegendKey' (bool) if true shows the legend values
     *      'showPercent' (bool) if true shows the percent values
     *      'showSeries' (bool) if true shows the series values
     *      'showTable' (bool) if true shows the table of values
     *      'showValue' (bool) if true shows the values inside the chart
     *      'stylesTitle' (array)
     *          'bold' (bool)
     *          'color' (string) FFFFFF, FF0000
     *          'font' (string)  Arial, Times New Roman ...
     *          'fontSize' (int) 8, 9, 10, ... size as drawing content (10 to 400000). 1420 as default
     *          'italic' (bool)
     *          'layout' (array)
     *              'x' (float) 0 < x < 1
     *              'y' (float) 0 < y < 1
     *      'tickLblPos' (mixed) tick label position (nextTo, high, low, none). If string, uses default values. If array, sets a value for each position
     *      'title' (string)
     *      'trendline' (array of trendlines). Compatible with line, bar and col 2D charts
     *          'color' (string) 0000FF
     *          'displayEquation' (bool) display equation on chart
     *          'displayRSquared' (bool) display R-squared value on chart
     *          'intercept' (float) set intercept
     *          'lineStyle' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *          'type' (string) 'exp', 'linear', 'log', 'poly', 'power', 'movingAvg'
     *          'typeOrder' (int) for poly and movingAvg types
     *      'vaxLabel' (string) vertical axis label
     *      'vaxLabelDisplay' (string) rotated, vertical, horizontal
     *      'vgrid'  (int) 0 (no grid) 1 (only major grid lines - default) 2 (only minor grid lines) 3 (both major and minor grid lines)
     *
     *  3D charts:
     *      'perspective' (int) 20, 30...
     *      'rotX' (int) 20, 30...
     *      'rotY' (int) 20, 30...
     *
     *  Bar and column charts:
     *      'gapWidth' (int) gap width
     *      'groupBar' (string) clustered, stacked, percentStacked
     *      'overlap' (int) overlap value
     *
     *  Line charts:
     *      'smooth' (mixed) enable smooth lines, line charts. '0' forces disabling it
     *      'symbol' (string) Line charts: none, dot, plus, square, star, triangle, x, diamond, circle and dash
     *      'symbolSize' (int) the size of the symbols (values 1 to 73)
     *
     *  Pie and doughnut charts:
     *      'explosion' (int) distance between the different values
     *      'holeSize' (int) size of the hole in doughnut type
     *
     *  Theme:
     *  'theme' (array):
     *      'chartArea' (array):
     *          'backgroundColor' (string)
     *      'gridLines' (array):
     *          'capType' (string)
     *          'color' (string): RGB
     *          'dashType' (string)
     *          'width' (int)
     *      'horizontalAxis' (array):
     *          'textBold' (bool)
     *          'textDirection' (string): 'horizontal', 'rotate90', 'rotate270'
     *          'textItalic' (bool)
     *          'textSize' (int): points
     *          'textUnderline' (string): DrawingML values such as 'none', 'sng', 'dash'
     *      'legendArea' (array):
     *          'backgroundColor' (string)
     *          'textBold' (bool)
     *          'textItalic' (bool)
     *          'textSize' (int): points
     *          'textUnderline' (string): DrawingML values such as 'none', 'sng', 'dash'
     *      'plotArea' (array):
     *          'backgroundColor' (string)
     *      'serDataLabels' (array): data labels options (bar, bubble, column, line, ofPie, pie and scatter charts)
     *          'fontStyles' (array):
     *              'bold' (bool)
     *              'color' (string) FFFFFF, FF0000...
     *              'font' (string) Arial, Times New Roman...
     *              'fontSize' (int) size as drawing content (100 to 400000) (100 = 1pt). 1420 as default
     *              'italic' (bool)
     *          'formatCode' (string)
     *          'position (string): center, insideEnd, insideBase, outsideEnd
     *          'showCategory' (int): 0, 1
     *          'showLegendKey' (int): 0, 1
     *          'showPercent' (int): 0, 1
     *          'showSeries' (int): 0, 1
     *          'showValue' (int): 0, 1
     *      'serRgbColors' (array): series colors
     *      'valueDataLabels' (array) data labels options (bar, bubble, column and line charts)
     *          'fontStyles' (array):
     *              'bold' (bool)
     *              'color' (string) FFFFFF, FF0000...
     *              'font' (string) Arial, Times New Roman...
     *              'fontSize' (int) size as drawing content (100 to 400000) (100 = 1pt). 1420 as default
     *              'italic' (bool)
     *          'position' (string): bottom, center, insideEnd, insideBase, left, outsideEnd, right, top, bestFit
     *          'showCategory' (int): 0, 1
     *          'showLegendKey' (int): 0, 1
     *          'showPercent' (int): 0, 1
     *          'showSeries' (int): 0, 1
     *          'showValue' (int): 0, 1
     *          'text' (string) this text hides other automatic values
     *      'valueRgbColors' (array): values colors
     *      'verticalAxis' (array):
     *          'textBold' (bool)
     *          'textDirection' (string): 'horizontal', 'rotate90', 'rotate270'
     *          'textItalic' (bool)
     *          'textSize' (int): points
     *          'textUnderline' (string): DrawingML values such as 'none', 'sng', 'dash'
     * @param array $options
     * @throws Exception chart type is not supported
     */
    public function addChart($chart, $position, $chartData, $chartStyles = array(), $options = array())
    {
        // extra options used to keep extra information
        $extraOptions = array();

        // check chart type
        if (!in_array($chart, array('area', 'area3D', 'bar', 'bar3D', 'bar3DCone', 'bar3DCylinder', 'bar3DPyramid', 'bubble', 'col', 'col3D', 'col3DCone', 'col3DCylinder', 'col3DPyramid', 'doughnut', 'line', 'line3D', 'ofPie', 'pie', 'pie3D', 'radar', 'scatter', 'surface'))) {
            PhppptxLogger::logger('Chart type is not supported.', 'fatal');
        }

        if (!file_exists(__DIR__ . '/ThemeCharts.php')) {
            unset($options['theme']);
        }

        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
        $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');

        $idChart = $this->generateUniqueId();
        $idChartExternalData = $this->generateUniqueId();
        $chartStyles['rId'] = $idChart;
        $chartStyles['rIdExternalData'] = $idChartExternalData;

        // create and add the new graphic frame
        $graphicFrameElement = new CreateGraphicFrame();
        $graphicFrameElement->addElementGraphicFrameChart($slideDOM, $position, $chartStyles, $options);

        // generate the new relationship
        $newRelationship = '<Relationship Id="rId'.$idChart.'" Target="../charts/chart'.$idChart.'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>';

        // generate content type if it does not exist yet
        $this->generateDEFAULT('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        // add Override
        $this->generateOVERRIDE('/ppt/charts/chart'.$idChart.'.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');

        // generate the XML chart
        $chartElement = new CreateChart();
        $elementsChart = $chartElement->createElementChart($chart, $chartData, $chartStyles, $options);

        if (isset($options['theme']) && is_array($options['theme']) && count($options['theme']) > 0) {
            $themeChart = new ThemeCharts();
            $elementsChart['chartXml'] = $themeChart->theme($elementsChart['chartXml'], $options['theme']);
        }

        // add the chart into the PPTX file
        $this->zipPptx->addContent('ppt/charts/chart'.$idChart.'.xml', $elementsChart['chartXml']);

        // add the external file
        $excelType = $elementsChart['chartType']->getXlsxType();
        $tempPath = TempDir::getTempDir();
        $this->tempFileXLSX[$idChart] = tempnam($tempPath, 'documentxlsx');
        $zipDocxExcel = $excelType->createChartXlsx($idChart, $chartData);
        $zipDocxExcel->savePptx($this->tempFileXLSX[$idChart], true);
        $this->zipPptx->addFile('ppt/embeddings/Microsoft_Excel_Worksheet' . $idChart . '.xlsx', $this->tempFileXLSX[$idChart] . '.pptx');

        // add the new relationship
        $this->generateRelationship($slideRelsDOM, $newRelationship);
        // generate and add the chart rels file
        $newRelationshipChart = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="../embeddings/Microsoft_Excel_Worksheet'.$idChart.'.xlsx" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"/></Relationships>';
        $this->zipPptx->addContent('ppt/charts/_rels/chart'.$idChart.'.xml.rels', $newRelationshipChart);

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());
        $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;
        $slideRelsDOM = null;

        PhppptxLogger::logger('Add chart.', 'info');
    }

    /**
     * Adds a comment
     *
     * @param string $content
     * @param string $author
     * @param array $position
     *      'coordinateX' (int) 10 as default
     *      'coordinateY' (int) 10 as default
     * @param array $options
     *      'date' (string) strtotime, 'now' as default
     * @throws Exception author doesn't exist
     */
    public function addComment($content, $author, $position = array(), $options = array())
    {
        // default values
        if (!isset($position['coordinateX'])) {
            $position['coordinateX'] = 10;
        }
        if (!isset($position['coordinateY'])) {
            $position['coordinateY'] = 10;
        }
        if (!isset($options['date'])) {
            $options['date'] = 'now';
        }

        // check if the comment author exists
        $commentAuthorsContents = $this->zipPptx->getContentByType('commentAuthors');
        if (count($commentAuthorsContents) == 0) {
            // commentAuthors file doesn't exist
            PhppptxLogger::logger('Comment author "' . $author . '" not found.', 'fatal');
        }
        $commentAuthorsDOM = $this->xmlUtilities->generateDomDocument($commentAuthorsContents[0]['content']);
        // check if the author exists
        $slideRelsXPath = new DOMXPath($commentAuthorsDOM);
        $slideRelsXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $nodesAuthor = $slideRelsXPath->query('//p:cmAuthorLst/p:cmAuthor[@name="'.$author.'"]');
        if ($nodesAuthor->length == 0) {
            PhppptxLogger::logger('Comment author "' . $author . '" not found.', 'fatal');
        }
        // keep author to be updated
        $nodeAuthor = $nodesAuthor->item(0);

        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];

        // get existing comment XML for the active slide if any exists. Generate a new one otherwise
        $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
        $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
        $slideRelsXPath = new DOMXPath($slideRelsDOM);
        $slideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $nodesSlideLayout = $slideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"]');
        if ($nodesSlideLayout->length > 0) {
            // a comment file exists, get it
            $targetComment = $nodesSlideLayout->item(0)->getAttribute('Target');

            // get the comment content
            $commentPath = str_replace('../', 'ppt/', $targetComment);
            $commentContentDOM = $this->zipPptx->getContent($commentPath, 'DOMDocument');
        } else {
            // generate and add a new comment file

            // generate a new Id for the comment
            $newId = $this->generateUniqueId();

            $commentPath = 'ppt/comments/comment'.$newId.'.xml';

            // generate the new relationship
            $newRelationship = '<Relationship Id="rId'.$newId.'" Target="../comments/comment'.$newId.'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" />';
            // add Override
            $this->generateOverride('/ppt/comments/comment'.$newId.'.xml', 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml');

            // add the new relationship
            $this->generateRelationship($slideRelsDOM, $newRelationship);

            // generate the comment content
            $commentContentDOM = $this->xmlUtilities->generateDomDocument(OOXMLResources::$skeletonComment);
        }

        // get all comment contents to get idx needed by the comment and the author
        $commentsContents = $this->zipPptx->getContentByType('comments');
        $currentIdx = 0;
        foreach ($commentsContents as $commentsContent) {
            if (!empty($commentsContent['content'])) {
                $commentsContentDOM = $this->xmlUtilities->generateDomDocument($commentsContent['content']);
                $nodesCm = $commentsContentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cm');
                foreach ($nodesCm as $nodeCm) {
                    if ((int)$nodeCm->getAttribute('idx') > $currentIdx) {
                        $currentIdx = (int)$nodeCm->getAttribute('idx');
                    }
                }
            }
        }

        // generate and add the new comment
        $commentAuthorId = $nodeAuthor->getAttribute('id');
        $commentDate = date("Y-m-d\TH:i:s.000", strtotime($options['date']));
        $commentIdx = $currentIdx + 1;

        $newCommentContent = '<p:cm xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" authorId="'.$commentAuthorId.'" dt="'.$commentDate.'" idx="'.$commentIdx.'"><p:pos x="'.$position['coordinateX'].'" y="'.$position['coordinateY'].'"/><p:text>'.$content.'</p:text></p:cm>';
        $newCommentFragment = $commentContentDOM->createDocumentFragment();
        $newCommentFragment->appendXML($newCommentContent);
        $commentContentDOM->documentElement->appendChild($newCommentFragment);

        // update the commentAuthors setting the new idx for the comment author
        $nodeAuthor->setAttribute('lastIdx', $commentIdx);

        // refresh contents
        $this->zipPptx->addContent($commentAuthorsContents[0]['path'], $commentAuthorsDOM->saveXML());
        $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());
        $this->zipPptx->addContent($commentPath, $commentContentDOM->saveXML());

        // free DOMDocument resources
        $commentAuthorsDOM = null;
        $slideRelsDOM = null;
        $commentContentDOM = null;

        PhppptxLogger::logger('Add comment.', 'info');
    }

    /**
     * Adds a comment author
     *
     * @param string $author
     * @param array $options
     *      'initials' (string) the same as $author if not set
     * @throws Exception duplicated author
     */
    public function addCommentAuthor($author, $options = array())
    {
        // default values
        if (!isset($options['initials'])) {
            $options['initials'] = $author;
        }

        // get existing comment authors XML if any exists. Generate a new one otherwise
        $commentAuthorsContents = $this->zipPptx->getContentByType('commentAuthors');
        if (count($commentAuthorsContents) == 0) {
            // commentAuthors file doesn't exist, generate and add a new one
            $commentAuthorsPath = 'ppt/commentAuthors.xml';
            // add Override
            $this->generateOverride('/ppt/commentAuthors.xml', 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml');

            // generate the commentAuthors content
            $commentAuthorsDOM = $this->xmlUtilities->generateDomDocument(OOXMLResources::$skeletonCommentAuthor);

            // add Relationship into the presentation
            $newId = $this->generateUniqueId();
            $newRelationship = '<Relationship Id="rId'.$newId.'" Target="commentAuthors.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"/>';
            $this->generateRelationship($this->pptxRelsPresentationDOM, $newRelationship);

            // refresh contents
            $this->zipPptx->addContent('ppt/_rels/presentation.xml.rels', $this->pptxRelsPresentationDOM->saveXML());
            $this->pptxRelsPresentationDOM = $this->zipPptx->getContent('ppt/_rels/presentation.xml.rels', 'DOMDocument');
        } else {
            // commentAuthors file exists
            $commentAuthorsPath = $commentAuthorsContents[0]['path'];

            $commentAuthorsDOM = $this->xmlUtilities->generateDomDocument($commentAuthorsContents[0]['content']);
        }

        // check if the author exists. Duplicated authors can't exist
        $slideRelsXPath = new DOMXPath($commentAuthorsDOM);
        $slideRelsXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $nodesAuthorName = $slideRelsXPath->query('//p:cmAuthorLst/p:cmAuthor[@name="'.$author.'"]');
        if ($nodesAuthorName->length > 0) {
            PhppptxLogger::logger('Duplicated author "' . $author . '".', 'fatal');
        }

        // get current author id and clrIdx values to generate a new one
        $authorId = 0;
        $authorClrIdx = -1;
        $nodesAuthors = $slideRelsXPath->query('//p:cmAuthorLst/p:cmAuthor');
        foreach ($nodesAuthors as $nodesAuthor) {
            if ($nodesAuthor->hasAttribute('id') && (int)$nodesAuthor->getAttribute('id') > $authorId) {
                $authorId = $nodesAuthor->getAttribute('id');
            }
            if ($nodesAuthor->hasAttribute('clrIdx') && (int)$nodesAuthor->getAttribute('clrIdx') > $authorClrIdx) {
                $authorClrIdx = $nodesAuthor->getAttribute('clrIdx');
            }
        }

        $authorId++;
        $authorClrIdx++;

        // generate and add the new comment author
        $newCommentAuthorContent = '<p:cmAuthor xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" clrIdx="'.$authorClrIdx.'" id="'.$authorId.'" initials="'.$options['initials'].'" lastIdx="0" name="'.$author.'"></p:cmAuthor>';
        $newCommentAuthorFragment = $commentAuthorsDOM->createDocumentFragment();
        $newCommentAuthorFragment->appendXML($newCommentAuthorContent);
        $commentAuthorsDOM->documentElement->appendChild($newCommentAuthorFragment);

        // refresh contents
        $this->zipPptx->addContent($commentAuthorsPath, $commentAuthorsDOM->saveXML());

        // free DOMDocument resources
        $commentAuthorsDOM = null;

        PhppptxLogger::logger('Add comment author.', 'info');
    }

    /**
     * Adds a footer in slide
     *
     * @param string $type Footer type: dateAndTime, slideNumber, textContents
     * @param array $options
     *      'applyToAll' (bool) if true apply to all slides. Default as false
     *      'contentStyles' (array) @see addText. Use with dateAndTime and slideNumber types. Using the textContents type, the PptxFragment must include the styles to be applied
     *      'dateAndTime' (string) set a fixed date and time. If not set, update automatically
     *      'dateAndTimeLanguage' (string) en-US, es-ES...
     *      'dateAndTimeType' (string) datetime, datetimeFigureOut, datetime1, datetime2, datetime3, datetime4, datetime5, datetime6, datetime7, datetime8, datetime9, datetime10, datetime11, datetime12, datetime13. Available when not using a fixed date and time
     *      'hideOn' (string) if set hide the footer on the slides that use the layout. Example: Title Slide
     *      'textContents' (PptxFragment) add a text content. Default as empty pararagraph
     * @throws Exception type not valid
     */
    public function addFooterSlide($type, $options = array())
    {
        $footerTypes = array('dateAndTime', 'slideNumber', 'textContents');
        if (!in_array($type, $footerTypes)) {
            PhppptxLogger::logger('Choose a valid footer type: dateAndTime, slideNumber or textContents.', 'fatal');
        }

        // default values
        if (!isset($options['applyToAll'])) {
            $options['applyToAll'] = false;
        }
        if (!isset($options['hideOn'])) {
            $options['hideOn'] = false;
        }

        // get the slides to be updated
        if (isset($options['applyToAll']) && $options['applyToAll']) {
            // all slides
            $slidesContents = $this->zipPptx->getSlides();

            // used to set slide number
            $indexSlide = 1;

            if (isset($options['hideOn']) & !empty($options['hideOn'])) {
                // remove slides that use Title Slide as layout
                $slidesContentsCleaned = array();
                foreach ($slidesContents as $slideContents) {
                    $slideContentDOM = $this->xmlUtilities->generateDomDocument($slideContents['content']);
                    $slideContentXPath = new DOMXPath($slideContentDOM);
                    $slideContentXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
                    $slideContentXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
                    $nodesPCSldTitleSlide = $slideContentXPath->query('//p:cSld[@name="'.$options['hideOn'].'"]');

                    if ($nodesPCSldTitleSlide->length > 0) {
                        continue;
                    }

                    $slidesContentsCleaned[] = $slideContents;

                    // free DOMDocument resources
                    $slideContentDOM = null;
                }

                $slidesContents = $slidesContentsCleaned;
            }
        } else {
            // active slide
            $slideContents = $this->zipPptx->getSlides();
            $slidesContents = array($slideContents[$this->activeSlide['position']]);
            // used to set slide number
            $indexSlide = $this->activeSlide['position'];
            $indexSlide += 1; // $this->activeSlide starts from 0 and slideNumber starts from 1
        }

        foreach ($slidesContents as $slidesContent) {
            // check if the slide has a p:sp tag with the chosen type
            // Order and type value: dateAndTime => dt, textContent => ftr, slideNumber => sldNum
            $slideDOM = $this->xmlUtilities->generateDomDocument($slidesContent['content']);
            $slideXPath = new DOMXPath($slideDOM);
            $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
            $slideXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

            // set footer type to be queried
            if ($type == 'dateAndTime') {
                $typePh = 'dt';
            } elseif ($type == 'slideNumber') {
                $typePh = 'sldNum';
            } else {
                $typePh = 'ftr';
            }

            $nodesFooterType = $slideXPath->query('//p:sp[.//p:nvPr/p:ph[@type='.$typePh.']]');
            if ($nodesFooterType->length == 0) {
                // the type doesn't exist in the slide. Get it from the slide layout and add the content to the slide
                $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $slidesContent['path']) . '.rels';
                $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
                $contentSlideRelsXPath = new DOMXPath($slideRelsDOM);
                $contentSlideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                $nodesSlideLayout = $contentSlideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"]');
                if ($nodesSlideLayout->length > 0) {
                    $targetSlideLayout = '';
                    if ($nodesSlideLayout->item(0)->hasAttribute('Target')) {
                        $targetSlideLayout = $nodesSlideLayout->item(0)->getAttribute('Target');
                    }
                    $nameSlideLayout = '';
                    if (!empty($targetSlideLayout)) {
                        $slideLayoutFilePath = str_replace('../', 'ppt/', $targetSlideLayout);
                        $contentSlideLayout = $this->zipPptx->getContent($slideLayoutFilePath);
                        if (!empty($contentSlideLayout)) {
                            $contentDOMSlideLayout = $this->xmlUtilities->generateDomDocument($contentSlideLayout);
                            $contentXPathSlideLayout = new DOMXPath($contentDOMSlideLayout);
                            $contentXPathSlideLayout->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
                            $nodesLayoutFooterType = $contentXPathSlideLayout->query('//p:sp[.//p:nvPr/p:ph[@type="'.$typePh.'"]]');
                            if ($nodesLayoutFooterType->length > 0) {
                                $newPSpFragment = $slideDOM->createDocumentFragment();
                                // add needed namespaces
                                $pSpContent = '<p:root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'.$nodesLayoutFooterType->item(0)->ownerDocument->saveXML($nodesLayoutFooterType->item(0)).'</p:root>';
                                $newPSpFragment->appendXML($pSpContent);
                                $nodesSpTree = $slideXPath->query('//p:spTree');
                                if ($nodesSpTree->length > 0) {
                                    $nodePSPFooterType = $nodesSpTree->item(0)->appendChild($newPSpFragment->firstChild->firstChild);
                                }
                            }

                            // free DOMDocument resources
                            $contentDOMSlideLayout = null;
                        }
                    }
                }

                // free DOMDocument resources
                $slideRelsDOM = null;
            } else {
                $nodePSPFooterType = $nodesFooterType->item(0);
            }

            if (isset($nodePSPFooterType)) {
                // text styles
                $textStyles = array();
                if (isset($options['contentStyles'])) {
                    // parse text and paragraph styles using CreateText
                    $text = new CreateText();
                    $options['contentStyles']['text'] = '';
                    $newContentText = $text->createElementText(array($options['contentStyles']), $options['contentStyles']);
                    $textStylesDOM = $this->xmlUtilities->generateDomDocument($newContentText);
                    $nodesPPr = $textStylesDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'pPr');
                    $nodesRPr = $textStylesDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'rPr');
                    if ($nodesPPr->length > 0) {
                        $textStyles['ppr'] = $nodesPPr->item(0)->ownerDocument->saveXML($nodesPPr->item(0));
                        if (strlen($textStyles['ppr']) < 10) {
                            // empty tag length, remove it
                            unset($textStyles['ppr']);
                        }
                    }
                    if ($nodesRPr->length > 0) {
                        $textStyles['rpr'] = $nodesRPr->item(0)->ownerDocument->saveXML($nodesRPr->item(0));
                        if (strlen($textStyles['rpr']) < 10) {
                            // empty tag length, remove it
                            unset($textStyles['rpr']);
                        }
                    }

                    // free DOMDocument resources
                    $textStylesDOM = null;
                }

                if ($type == 'dateAndTime') {
                    $nodesT = $slideXPath->query('.//p:txBody//a:p//a:t', $nodePSPFooterType);
                    if ($nodesT->length > 0) {
                        // update date content
                        $dateAndTime = date('Y/m/d');
                        if (isset($options['dateAndTime'])) {
                            // fixed date time
                            $dateAndTime = $options['dateAndTime'];
                        }
                        $nodesT->item(0)->nodeValue = $dateAndTime;
                        // update date options
                        if (isset($options['dateAndTimeLanguage'])) {
                            // date and time language
                            $nodesRPR = $slideXPath->query('.//p:txBody//a:p//a:rPr', $nodePSPFooterType);
                            foreach ($nodesRPR as $nodeRPR) {
                                $nodeRPR->setAttribute('lang', $options['dateAndTimeLanguage']);
                            }
                        }
                        if (isset($options['dateAndTimeType'])) {
                            // date and time type
                            $nodesFld = $slideXPath->query('.//p:txBody//a:p//a:fld', $nodePSPFooterType);
                            foreach ($nodesFld as $nodeFld) {
                                $nodeFld->setAttribute('type', $options['dateAndTimeType']);
                            }
                        }
                        if (isset($options['dateAndTime'])) {
                            // set fixed date and time tag
                            $nodesFld = $slideXPath->query('.//p:txBody//a:p//a:fld', $nodePSPFooterType);
                            $fldNodesToBeRemoved = array();
                            foreach ($nodesFld as $nodeFld) {
                                // replace a:fld with a:r
                                $fldNodesToBeRemoved[] = $nodeFld;
                                $newRNodeContent = '<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
                                foreach ($nodeFld->childNodes as $nodeFldChild) {
                                    $newRNodeContent .= $nodeFldChild->ownerDocument->saveXML($nodeFldChild);
                                }
                                $newRNodeContent .= '</a:r>';
                                $newRragment = $nodeFld->ownerDocument->createDocumentFragment();
                                $newRragment->appendXML($newRNodeContent);
                                $nodeFld->parentNode->insertBefore($newRragment, $nodeFld->nextSibling);
                            }
                            // remove old nodes
                            foreach ($fldNodesToBeRemoved as $fldNodeToBeRemoved) {
                                $fldNodeToBeRemoved->parentNode->removeChild($fldNodeToBeRemoved);
                            }
                        }
                    }
                } elseif ($type == 'slideNumber') {
                    $nodesT = $slideXPath->query('.//p:txBody//a:p//a:t', $nodePSPFooterType);
                    if ($nodesT->length > 0) {
                        $nodesT->item(0)->nodeValue = $indexSlide;
                    }
                } elseif ($type == 'textContents') {
                    $nodesP = $slideXPath->query('.//p:txBody//a:p', $nodePSPFooterType);
                    foreach ($nodesP as $nodeP) {
                        $nodeP->parentNode->removeChild($nodeP);
                    }
                    $nodesTxBody = $slideXPath->query('.//p:txBody', $nodePSPFooterType);
                    if ($nodesTxBody->length > 0) {
                        $textContents = '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />';
                        if (isset($options['textContents']) && $options['textContents'] instanceof PptxFragment) {
                            $textContents = (string)$options['textContents'];
                        }

                        $newPFragment = $slideDOM->createDocumentFragment();
                        $newPFragment->appendXML($textContents);
                        $nodesTxBody->item(0)->appendChild($newPFragment);
                    }
                }

                // apply content styles to dateAndTime and slideNumber. textContents styles must be applied in the PptxFragment
                if (count($textStyles) > 0 && ($type == 'dateAndTime' || $type == 'slideNumber')) {
                    $nodesP = $slideXPath->query('.//p:txBody//a:p', $nodePSPFooterType);
                    if (isset($textStyles['ppr'])) {
                        foreach ($nodesP as $nodeP) {
                            $nodesPPr = $slideXPath->query('.//a:pPr', $nodeP);
                            $newAPPrNodeContent = '<a:root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'.$textStyles['ppr'].'</a:root>';
                            $newAPPrFragment = $nodeP->ownerDocument->createDocumentFragment();
                            $newAPPrFragment->appendXML($newAPPrNodeContent);
                            $newAPPrNode = $newAPPrFragment->firstChild->firstChild;
                            if ($nodesPPr->length > 0) {

                                $pPrNodesToBeRemoved = array();

                                foreach ($nodesPPr as $nodePPr) {
                                    $pPrNodesToBeRemoved[] = $nodePPr;

                                    $nodePPr->parentNode->insertBefore($newAPPrNode, $nodePPr->nextSibling);
                                }

                                // remove old nodes
                                foreach ($pPrNodesToBeRemoved as $pPrNodeToBeRemoved) {
                                    $pPrNodeToBeRemoved->parentNode->removeChild($pPrNodeToBeRemoved);
                                }
                            } else {
                                // no pPr, add the new one as first child
                                $nodeP->insertBefore($newAPPrNode, $nodeP->firstChild);
                            }
                        }
                    }
                    if (isset($textStyles['rpr'])) {
                        foreach ($nodesP as $nodeP) {
                            $nodesRPr = $slideXPath->query('.//a:rPr', $nodeP);
                            $newARPrNodeContent = '<a:root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'.$textStyles['rpr'].'</a:root>';
                            $newARPrFragment = $nodeP->ownerDocument->createDocumentFragment();
                            $newARPrFragment->appendXML($newARPrNodeContent);
                            $newARPrNode = $newARPrFragment->firstChild->firstChild;
                            if ($nodesRPr->length > 0) {
                                $rPrNodesToBeRemoved = array();

                                foreach ($nodesRPr as $nodeRPr) {
                                    $rPrNodesToBeRemoved[] = $nodeRPr;

                                    $nodeRPr->parentNode->insertBefore($newARPrNode, $nodeRPr->nextSibling);
                                }

                                // remove old nodes
                                foreach ($rPrNodesToBeRemoved as $rPrNodeToBeRemoved) {
                                    $rPrNodeToBeRemoved->parentNode->removeChild($rPrNodeToBeRemoved);
                                }
                            } else {
                                // no pPr, add the new one as first child of a:fld and a:r
                                $nodesFlrR = $slideXPath->query('.//a:fld | .//a:r', $nodeP);
                                foreach ($nodesFlrR as $nodeFlrR) {
                                    $nodeFlrR->insertBefore($newARPrNode, $nodeFlrR->firstChild);
                                }
                            }
                        }
                    }
                }
            }

            $this->zipPptx->addContent($slidesContent['path'], $slideDOM->saveXML());

            $indexSlide++;

            // free DOMDocument resources
            $slideDOM = null;
        }

        PhppptxLogger::logger('Add footer in slide.', 'info');
    }

    /**
     * Adds HTML
     *
     * @param string $html HTML to add
     * @param array $position
     *      'placeholder' (array) the content is added in a placeholder of the layout. One of the following options can be used to get the text box
     *          'name' (string) name
     *          'descr' (string) alt text (descr) value
     *          'position' (int) position by order. 0 is the first order position
     *          'type' (string) title (Title), body (Body), ctrTitle (Centered Title), subTitle (Subtitle)
     *      'new' (array) a new position is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     *          'textBoxStyles' (array) @see addTextBox
     * @param array $options
     *      'disableWrapValue' (bool) if true disable using a wrap value with Tidy. Default as false
     *      'forceNotTidy' (bool) if true, avoid using Tidy. Only recommended if Tidy can't be installed. Default as false
     *      'insertMode' (string) append, replace. If a content exists in the text box, handle how new contents are added. Default as append
     *      'parseCSSVars' (bool) parse CSS variables. Default as false
     *      'useHtmlExtended' (bool) use HTML Extended and CSS Extended. Default as false. Premium licenses
     * @throws Exception PHP Tidy is not enabled
     * @throws Exception position not valid
     */
    public function addHtml($html, $position, $options = array())
    {
        // default options
        if (!isset($options['baseURL'])) {
            $options['baseURL'] = '';
        }
        if (!isset($options['disableWrapValue'])) {
            $options['disableWrapValue'] = false;
        }
        if (!isset($options['forceNotTidy'])) {
            $options['forceNotTidy'] = false;
        }
        if (!isset($options['insertMode'])) {
            $options['insertMode'] = 'append';
        }
        if (!isset($options['parseAnchors'])) {
            $options['parseAnchors'] = false;
        }
        if (!isset($options['parseCSSVars'])) {
            $options['parseCSSVars'] = false;
        }
        if (!isset($options['parseDivs'])) {
            $options['parseDivs'] = false;
        }
        if (!isset($options['useHtmlExtended'])) {
            $options['useHtmlExtended'] = false;
        }

        // keep the position to be added in HTML Extended tags
        $options['positionHtml'] = $position;

        // set a default value if empty to avoid a PHP fatal error
        if (empty($html)) {
            $html = '<html><body><p>&nbsp;</p></body></html>';
        }

        if (!extension_loaded('tidy') && !$options['forceNotTidy']) {
            PhppptxLogger::logger('Install and enable Tidy for PHP (https://php.net/manual/en/book.tidy.php) to transform HTML to PPTX. If PHP Tidy can\'t be installed, enable the forceNotTidy option.', 'fatal');
        }

        if (!($this instanceof PptxFragment)) {
            // do the transformation
            $htmlElement = new CreateHtml($this);
            $newContentHtml = $htmlElement->createElementHtml($html, $options);

            // get the internal active slide after adding the HTML to do not override contents added in the slide DOM by HTML Extended tags
            $slideContents = $this->zipPptx->getSlides();
            $activeSlideContent = $slideContents[$this->activeSlide['position']];
            $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
            $slideXPath = new DOMXPath($slideDOM);
            $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

            if (!empty($newContentHtml)) {
                // handle external relationships such as hyperlinks
                $externalRelationships = $htmlElement->getExternalRelationships();
                if (count($externalRelationships) > 0) {
                    $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
                    $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
                    $this->addExternalRelationships($externalRelationships, $activeSlideContent['path'], $slideRelsDOM);

                    // refresh contents
                    $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

                    // free DOMDocument resources
                    $slideRelsDOM = null;
                }

                // get the box to add the new content
                $nodePSp = $this->getPspBox($position, $slideDOM, $slideXPath, $options);

                // insert the new content
                $nodesTxBody = $nodePSp->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'txBody');
                if ($nodesTxBody->length > 0) {
                    // clean existing contents
                    $this->cleanPspBox($nodesTxBody->item(0), $slideXPath, $options);

                    // append the new contents
                    $newTextFragment = $slideDOM->createDocumentFragment();
                    $newTextFragment->appendXML($newContentHtml);
                    $nodesTxBody->item(0)->appendChild($newTextFragment);
                }

                // refresh contents
                $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());
            }

            // free DOMDocument resources
            $slideDOM = null;

            PhppptxLogger::logger('Add HTML content.', 'info');
        } else {
            // do the transformation
            $htmlElement = new CreateHtml($this);
            $this->pptxML .= $htmlElement->createElementHtml($html, $options);

            // handle external relationships such as hyperlinks
            $externalRelationships = $htmlElement->getExternalRelationships();
            foreach ($externalRelationships as $externalRelationship) {
                $this->addExternalRelationshipFragment($externalRelationship);
            }
        }
    }

    /**
     * Adds an image
     *
     * @param mixed $image Image path, base64, stream or GdImage. Image formats: png, jpg, jpeg, gif, bmp, webp
     * @param array $position
     *      'new' (array) a new shape is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $imageStyles
     *      'border' (array)
     *          'dash' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *          'color' (string) default as 000000
     *          'width' (int) EMUs (English Metric Unit). Default as 12700 (1pt)
     *      'descr' (string) set a descr value
     *      'hyperlink' (string) hyperlink. External, bookmark (#firstslide, #lastslide, #nextslide, #previousslide) or slide (#slide + position)
     *      'rotation' (int)
     * @param array $options
     *      'mime' (string) forces a mime
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     * @throws Exception mime option is not set and getimagesizefromstring is not available
     * @throws Exception size not valid
     * @throws Exception hyperlink slide position not valid
     * @throws Exception position not valid
     */
    public function addImage($image, $position, $imageStyles = array(), $options = array())
    {
        // get image information
        $imageInformation = new ImageUtilities();
        $imageContents = $imageInformation->returnImageContents($image, $options);
        $imageStyles['contents'] = $imageContents;

        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
        $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');

        $idImage = $this->generateUniqueId();
        $imageStyles['rId'] = $idImage;

        // create and add the new pic
        $imageElement = new CreatePic();
        $imageElement->addElementImage($slideDOM, $position, $imageStyles, $options);
        $externalRelationships = $imageElement->getExternalRelationships();

        // handle external relationships such as hyperlinks
        $this->addExternalRelationships($externalRelationships, $activeSlideContent['path'], $slideRelsDOM);

        // generate the new relationship
        $newRelationship = '<Relationship Id="rId'.$idImage.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img'.$idImage.'.'.$imageContents['extension'].'" />';
        // generate content type if it does not exist yet
        $this->generateDefault($imageContents['extension'], 'image/' . $imageContents['extension']);

        // copy the image with the new name
        $this->zipPptx->addContent('ppt/media/img'.$idImage.'.'.$imageContents['extension'], $imageContents['content']);

        // add the new relationship
        $this->generateRelationship($slideRelsDOM, $newRelationship);

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());
        $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;
        $slideRelsDOM = null;

        PhppptxLogger::logger('Add image.', 'info');
    }

    /**
     * Adds a link
     *
     * @access public
     * @param string $link external, bookmark (#firstslide, #lastslide, #nextslide, #previousslide) or slide (#slide with position)
     * @param string $linkText Text Content
     * @param array $position
     *      'placeholder' (array) the content is added in a placeholder of the layout. One of the following options can be used to get the text box
     *          'name' (string) name
     *          'descr' (string) alt text (descr) value
     *          'position' (int) position by order. 0 is the first order position
     *          'type' (string) title (Title), body (Body), ctrTitle (Centered Title), subTitle (Subtitle)
     *      'new' (array) a new position is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     *          'textBoxStyles' (array) @see addTextBox
     * @param array $linkStyles @see addText
     * @param array $paragraphStyles @see addText
     * @param array $options
     *      'insertMode' (string) append, replace. If a content exists in the text box, handle how new contents are added. Default as append
     * @throws Exception position not valid
     * @throws Exception hyperlink slide position not valid
     */
    public function addLink($link, $linkText, $position, $linkStyles = array(), $paragraphStyles = array(), $options = array())
    {
        $linkStyles['hyperlink'] = $link;
        $linkStyles['text'] = $linkText;
        $this->addText($linkStyles, $position, $paragraphStyles, $options);

        PhppptxLogger::logger('Add link.', 'info');
    }

    /**
     * Adds a list
     *
     * @access public
     * @param array $contents array of contents or PptxFragments
     *      'text' (string|array) @see addText
     *      'bold' (bool)
     *      'characterSpacing' (int)
     *      'color' (string) HEX color
     *      'font' (string) Arial, Times New Roman ...
     *      'fontSize' (int) 8, 9, 10, 11 ...
     *      'highlight' (string) HEX color
     *      'italic' (bool)
     *      'lang' (string)
     *      'strikethrough' (bool)
     *      'underline' (string) single
     * @param array $position
     *      'placeholder' (array) the content is added in a placeholder of the layout. One of the following options can be used to get the text box
     *          'name' (string) name
     *          'descr' (string) alt text (descr) value
     *          'position' (int) position by order. 0 is the first order position
     *          'type' (string) title (Title), body (Body), ctrTitle (Centered Title), subTitle (Subtitle)
     *      'new' (array) a new position is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     *          'textBoxStyles' (array) @see addTextBox
     * @param array $listStyles for each list level
     *      'color' (string)
     *      'font' (string)
     *      'indent' (int) EMUs (English Metric Unit) default as -250000
     *      'marginLeft' (int) EMUs (English Metric Unit)
     *      'marginRight' (int) EMUs (English Metric Unit)
     *      'size' (int) % of text
     *      'startAt' (int)
     *      'type' (string) filledRoundBullet, hollowRoundBullet, filledSquareBullet, hollowSquareBullet, starBullet, arrowBullet, checkmarkBullet, decimal, romanUpperCase, romanLowerCase, alphaUpperCase, alphaLowerCase. Other types: alphaLcParenBoth, alphaUcParenBoth, alphaLcParenR, alphaUcParenR, alphaLcPeriod, alphaUcPeriod, arabicParenBoth, arabicParenR, arabicPeriod, arabicPlain, romanLcParenBoth, romanUcParenBoth, romanLcParenR, romanUcParenR, romanLcPeriod, romanUcPeriod, circleNumDbPlain, circleNumWdBlackPlain, circleNumWdWhitePlain, arabicDbPeriod, arabicDbPlain, ea1ChsPeriod, ea1ChsPlain, ea1ChtPeriod, ea1ChtPlain, ea1JpnChsDbPeriod, ea1JpnKorPlain, ea1JpnKorPeriod, arabic1Minus, arabic2Minus, hebrew2Minus, thaiAlphaPeriod, thaiAlphaParenR, thaiAlphaParenBoth, thaiNumPeriod, thaiNumParenR, thaiNumParenBoth, hindiAlphaPeriod, hindiNumPeriod, hindiNumParenR, hindiAlpha1Period
     * @param array $options
     *      'insertMode' (string) append, replace. If a content exists in the text box, handle how new contents are added. Default as append
     * @throws Exception position not valid
     */
    public function addList($contents, $position, $listStyles = array(), $options = array())
    {
        // default values
        if (!isset($options['insertMode'])) {
            $options['insertMode'] = 'append';
        }

        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // get the box to add the new content
        $nodePSp = $this->getPspBox($position, $slideDOM, $slideXPath, $options);

        $nodesTxBody = $nodePSp->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'txBody');
        if ($nodesTxBody->length > 0) {
            // clean existing contents
            $this->cleanPspBox($nodesTxBody->item(0), $slideXPath, $options);

            // insert the list values in a recursive way
            $this->insertListValues($slideDOM, $activeSlideContent, $nodesTxBody->item(0), $contents, $listStyles, 0);
        }

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;

        PhppptxLogger::logger('Add list.', 'info');
    }

    /**
     * Adds a macro from a PPTX
     *
     * @access public
     * @param string $source Path to a file with macro
     * @throws Exception the file can't be opened
     * @throws Exception macro not found
     */
    public function addMacroFromPptx($source)
    {
        $pptxMacro = new ZipArchive();
        if ($pptxMacro->open($source) !== TRUE) {
            PhppptxLogger::logger('Error while trying to open \'' . $source . '\' as PPTM.', 'fatal');
        }

        // generate new rels and ContentTypes
        $this->generateOverride('/ppt/vbaProject.bin', 'application/vnd.ms-office.vbaProject');

        // add Relationship if no previous vbaProject.bin Relationship exists
        $pptxRelsPresentationDOM = new DOMXPath($this->pptxRelsPresentationDOM);
        $pptxRelsPresentationDOM->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $nodesRelationshipVbaProject = $pptxRelsPresentationDOM->query('//xmlns:Relationships/xmlns:Relationship[@Target="vbaProject.bin"]');
        if ($nodesRelationshipVbaProject->length == 0) {
            // add Relationship into the presentation
            $newId = $this->generateUniqueId();
            $newRelationship = '<Relationship Id="rId'.$newId.'" Target="vbaProject.bin" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"/>';
            $this->generateRelationship($this->pptxRelsPresentationDOM, $newRelationship);

            // refresh contents
            $this->zipPptx->addContent('ppt/_rels/presentation.xml.rels', $this->pptxRelsPresentationDOM->saveXML());
            $this->pptxRelsPresentationDOM = $this->zipPptx->getContent('ppt/_rels/presentation.xml.rels', 'DOMDocument');
        }

        // set /ppt/presentation.xml Relationship as macro Override type
        $pptxContentTypesXPath = new DOMXPath($this->pptxContentTypesDOM);
        $pptxContentTypesXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');
        $nodesOverrideMain = $pptxContentTypesXPath->query('//xmlns:Types/xmlns:Override[@ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"]');
        if ($nodesOverrideMain->length > 0) {
            $nodesOverrideMain->item(0)->setAttribute('ContentType', 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml');

            // refresh contents
            $this->zipPptx->addContent('[Content_Types].xml', $this->pptxContentTypesDOM->saveXML());
            $this->pptxContentTypesDOM = $this->zipPptx->getContent('[Content_Types].xml', 'DOMDocument');
        }

        // get and copy the contents of vbaData
        $vbaProjectBinFile = $pptxMacro->getFromName('ppt/vbaProject.bin');
        if (!$vbaProjectBinFile) {
            PhppptxLogger::logger('Macro not found.', 'fatal');
        }
        $this->zipPptx->addContent('ppt/vbaProject.bin', $vbaProjectBinFile);
        $pptxMacro->close();

        $this->isMacro = true;
        $this->zipPptx->setFileType('pptm');

        PhppptxLogger::logger('Add macro file.', 'info');
    }

    /**
     * Adds a math equation
     *
     * @access public
     * @param string $equation OMML equation string or MathML
     * @param string $type Type of equation: omml, mathml
     * @param array $position
     *      'placeholder' (array) the content is added in a placeholder of the layout. One of the following options can be used to get the text box
     *          'name' (string) name
     *          'descr' (string) alt text (descr) value
     *          'position' (int) position by order. 0 is the first order position
     *          'type' (string) title (Title), body (Body), ctrTitle (Centered Title), subTitle (Subtitle)
     *      'new' (array) a new position is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     *          'textBoxStyles' (array) @see addTextBox
     * @param array $options
     *      'align' (string) left, center, right
     *      'color' (string) FFFFFF, FF0000...
     *      'fontSize' (int) 8, 9, 10...
     * @throws Exception not valid type
     * @throws Exception position not valid
     */
    public function addMathEquation($equation, $type, $position, $options = array())
    {
        // default options
        if (!isset($options['insertMode'])) {
            $options['insertMode'] = 'append';
        }

        if ($type != 'omml' && $type != 'mathml') {
            PhppptxLogger::logger('Choose a valid type of equation: omml or mathml.', 'fatal');
        }

        if (!($this instanceof PptxFragment)) {
            // get the internal active slide
            $slideContents = $this->zipPptx->getSlides();
            $activeSlideContent = $slideContents[$this->activeSlide['position']];
            $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
            $slideXPath = new DOMXPath($slideDOM);
            $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

            $options['isPptxFragment'] = false;

            // create the math tags
            $math = new CreateMath();
            $newContentMath = $math->createElementMath($equation, $type, $options);

            // get the box to add the new content
            $nodePSp = $this->getPspBox($position, $slideDOM, $slideXPath, $options);

            // insert the new content
            $nodesTxBody = $nodePSp->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'txBody');
            if ($nodesTxBody->length > 0) {
                // clean existing contents
                $this->cleanPspBox($nodesTxBody->item(0), $slideXPath, $options);

                // append the new contents
                $newMathFragment = $slideDOM->createDocumentFragment();
                $newMathFragment->appendXML($newContentMath);
                $nodesTxBody->item(0)->appendChild($newMathFragment);
            }

            // refresh contents
            $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());

            // free DOMDocument resources
            $slideDOM = null;

            PhppptxLogger::logger('Add math equation.', 'info');
        } else {
            $options['isPptxFragment'] = true;

            // create the math tags
            $math = new CreateMath();
            $this->pptxML .= $math->createElementMath($equation, $type, $options);
        }
    }

    /**
     * Adds notes
     *
     * @param string|array $content @see addText
     * @param array $paragraphStyles @see addText
     * @param array $options
     *      'insertMode' (string) append, replace. If a note exists in the slide, handle how new notes are added. Default as append
     * @throws Exception there must be at least one slide
     */
    public function addNotes($content, $paragraphStyles = array(), $options = array())
    {
        // default values
        if (!isset($options['insertMode'])) {
            $options['insertMode'] = 'append';
        }

        // check if a notesMaster exists
        $notesMasterIdLstTags = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'notesMasterIdLst');
        if ($notesMasterIdLstTags->length == 0) {
            // notesMaster file doesn't exist, create and add a new one

            // generate a new ID for the notesMaster
            $newIdNotesMasterIdLst = $this->generateUniqueId();
            $sldMasterIdLstTags = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldMasterIdLst');
            if ($sldMasterIdLstTags->length > 0) {
                // insert p:notesMasterIdLst tag in the presentation
                $newNotesMasterIdLstFragment = $this->pptxPresentationDOM->createDocumentFragment();
                $newNotesMasterIdLstFragment->appendXML('<p:notesMasterIdLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:notesMasterId r:id="rId'.$newIdNotesMasterIdLst.'"/></p:notesMasterIdLst>');
                $sldMasterIdLstTags->item(0)->parentNode->insertBefore($newNotesMasterIdLstFragment, $sldMasterIdLstTags->item(0)->nextSibling);

                // add Relationship
                $newRelationship = '<Relationship Id="rId'.$newIdNotesMasterIdLst.'" Target="notesMasters/notesMaster'.$newIdNotesMasterIdLst.'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"/>';
                $this->generateRelationship($this->pptxRelsPresentationDOM, $newRelationship);

                // add Override
                $this->generateOverride('/ppt/notesMasters/notesMaster'.$newIdNotesMasterIdLst.'.xml', 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml');

                // add the new notesMaster file
                $this->zipPptx->addContent('ppt/notesMasters/notesMaster'.$newIdNotesMasterIdLst.'.xml', OOXMLResources::$skeletonNotesMaster);

                // add notesMaster rels
                $themesContents = $this->zipPptx->getContentByType('themes');
                if (count($themesContents) > 0) {
                    // generate a new ID for the relationship
                    $newIdRelationship = $this->generateUniqueId();

                    // add a new theme. Required by notesMaster
                    $this->generateOverride('/ppt/theme/theme'.$newIdRelationship.'.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');
                    $this->zipPptx->addContent('ppt/theme/theme'.$newIdRelationship.'.xml', $themesContents[0]['content']);

                    $this->zipPptx->addContent('ppt/notesMasters/_rels/notesMaster'.$newIdNotesMasterIdLst.'.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId'.$newIdRelationship.'" Target="../theme/theme'.$newIdRelationship.'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"/></Relationships>');
                }

                // refresh contents
                $this->zipPptx->addContent('ppt/presentation.xml', $this->pptxPresentationDOM->saveXML());
                $this->zipPptx->addContent('ppt/_rels/presentation.xml.rels', $this->pptxRelsPresentationDOM->saveXML());
                $this->pptxPresentationDOM = $this->zipPptx->getContent('ppt/presentation.xml', 'DOMDocument');
                $this->pptxRelsPresentationDOM = $this->zipPptx->getContent('ppt/_rels/presentation.xml.rels', 'DOMDocument');
            } else {
                PhppptxLogger::logger('There must be at least one slide.', 'fatal');
            }
        }

        $notesMasterContents = $this->zipPptx->getContentByType('notesMasters');
        if (count($notesMasterContents)> 0) {
            // get the internal active slide
            $slideContents = $this->zipPptx->getSlides();
            $activeSlideContent = $slideContents[$this->activeSlide['position']];

            // check if a notes exists in the active slide
            $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
            $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
            $slideRelsXPath = new DOMXPath($slideRelsDOM);
            $slideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
            $nodesSlideNotes = $slideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"]');
            if ($nodesSlideNotes->length > 0) {
                // a notes file exists, get it
                $targetNotes = $nodesSlideNotes->item(0)->getAttribute('Target');

                // get the notes content
                $notesPath = str_replace('../', 'ppt/', $targetNotes);
                $notesContentDOM = $this->zipPptx->getContent($notesPath, 'DOMDocument');
            } else {
                // generate and add a new notes file

                // generate a new Id for the notes
                $newNotesId = $this->generateUniqueId();

                $notesPath = 'ppt/notesSlides/notesSlide'.$newNotesId.'.xml';

                // generate the new relationship
                $newRelationship = '<Relationship Id="rId'.$newNotesId.'" Target="../notesSlides/notesSlide'.$newNotesId.'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" />';
                // add Override
                $this->generateOverride('/ppt/notesSlides/notesSlide'.$newNotesId.'.xml', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');

                // add the new relationship
                $this->generateRelationship($slideRelsDOM, $newRelationship);

                // add notesSlide rels
                $notesMastersPath = str_replace('ppt/notesMasters/', '../notesMasters/', $notesMasterContents[0]['path']);
                $notesSlidesPath = str_replace('ppt/slides/', '../slides/', $activeSlideContent['path']);
                // generate new IDs for the relationships
                $newIdRelationshipNotesSlides = $this->generateUniqueId();
                $newIdRelationshipNotesMaster = $this->generateUniqueId();
                $this->zipPptx->addContent('ppt/notesSlides/_rels/notesSlides'.$newNotesId.'.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId'.$newIdRelationshipNotesSlides.'" Target="'.$notesSlidesPath.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/><Relationship Id="rId'.$newIdRelationshipNotesMaster.'" Target="'.$notesMastersPath.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"/></Relationships>');

                // generate the notes content
                $notesContentDOM = $this->xmlUtilities->generateDomDocument(OOXMLResources::$skeletonNotes);
            }

            // add notes content
            // allow string as $contents instead of an array. Transform string to array
            if (!is_array($content)) {
                $contentsNormalized = array();
                $contentsNormalized['text'] = $content;
                $content = $contentsNormalized;
            }
            // if not using a subarray, generate it
            if (isset($content['text'])) {
                $content = array($content);
            }
            $text = new CreateText();
            $notesContent = $text->createElementText($content, $paragraphStyles);

            // generate random values to set the notes
            $noteId = mt_rand(999, 999999);
            $noteName = 'Notes Placeholder ' . $noteId;
            // check if the notesMaster contents has a preset p:sp to add the text contents
            $notesMasterDOM = $this->xmlUtilities->generateDomDocument($notesMasterContents[0]['content']);
            $notesMasterXPath = new DOMXPath($notesMasterDOM);
            $notesMasterXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
            $nodesPSpBody = $notesMasterXPath->query('//p:sp[.//p:nvPr/p:ph[@type="body"]]');
            if ($nodesPSpBody->length > 0) {
                $nodesCNvPr = $notesMasterXPath->query('.//p:cNvPr', $nodesPSpBody->item(0));
                if ($nodesCNvPr->length > 0) {
                    if ($nodesCNvPr->item(0)->hasAttribute('id')) {
                        $noteId = $nodesCNvPr->item(0)->getAttribute('id');
                    }
                    if ($nodesCNvPr->item(0)->hasAttribute('name')) {
                        $noteName = $nodesCNvPr->item(0)->getAttribute('name');
                    }
                }
            }

            // handle how the note is added
            if ($nodesPSpBody->length == 0) {
                // add the new note as a new shape using a default skeleton
                $notesContent = '<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:nvSpPr><p:cNvPr id="'.$noteId.'" name="'.$noteName.'"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1" type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'.$notesContent.'</p:txBody></p:sp>';
                $nodesSpTree = $notesContentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'spTree');
                if ($nodesSpTree->length > 0) {
                    $notesFragment = $notesContentDOM->createDocumentFragment();
                    $notesFragment->appendXML($notesContent);
                    $nodesSpTree->item(0)->appendChild($notesFragment);
                }
            } else {
                // add the new note in the existing shape

                // check if the notes file has a shape to add the new notes
                $notesContentXPath = new DOMXPath($notesContentDOM);
                $notesContentXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
                $notesContentXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

                $nodesContentPSpBody = $notesContentXPath->query('//p:sp[.//p:nvPr/p:ph[@type="body"]]');
                if ($nodesContentPSpBody->length == 0) {
                    // add the new note as a new shape using the notesMaster shape
                    $nodesP = $notesMasterXPath->query('.//a:p', $nodesPSpBody->item(0));
                    foreach ($nodesP as $nodeP) {
                        $nodeP->parentNode->removeChild($nodeP);
                    }
                    $nodesSpTree = $notesContentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'spTree');
                    if ($nodesSpTree->length > 0) {
                        $notesContent = str_replace('</p:txBody>', $notesContent . '</p:txBody>', $nodesPSpBody->item(0)->ownerDocument->saveXML($nodesPSpBody->item(0)));
                        $notesContent = str_replace('<p:sp>', '<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">', $notesContent);
                        $notesFragment = $notesContentDOM->createDocumentFragment();
                        $notesFragment->appendXML($notesContent);
                        $nodesSpTree->item(0)->appendChild($notesFragment);
                    }
                } else {
                    // add the new note using the notes shape

                    if ($options['insertMode'] == 'replace') {
                        // remove existing notes and add new notes
                        $nodesP = $notesContentXPath->query('.//a:p', $nodesContentPSpBody->item(0));
                        foreach ($nodesP as $nodeP) {
                            $nodeP->parentNode->removeChild($nodeP);
                        }
                    }

                    $nodesBody = $notesContentXPath->query('.//p:txBody', $nodesContentPSpBody->item(0));
                    if ($nodesBody->length > 0) {
                        $notesFragment = $notesContentDOM->createDocumentFragment();
                        $notesFragment->appendXML($notesContent);
                        $nodesBody->item(0)->appendChild($notesFragment);
                    }
                }
            }

            // refresh contents
            $this->zipPptx->addContent($notesPath, $notesContentDOM->saveXML());
            $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

            // free DOMDocument resources
            $notesContentDOM = null;
            $notesMasterContentDOM = null;
        }

        // free DOMDocument resources
        $slideRelsDOM = null;

        PhppptxLogger::logger('Add notes.', 'info');
    }

    /**
     * Adds properties
     *
     * @access public
     * @param array $values
     *      'category' (string)
     *      'Company' (string)
     *      'contentStatus' (string)
     *      'created' (string) W3CDTF without time zone
     *      'creator' (string)
     *      'custom' (array)
     *          'name' (array) 'type' => 'value'
     *      'description' (string)
     *      'keywords' (string)
     *      'lastModifiedBy' (string)
     *      'Manager' (string)
     *      'modified' (string) W3CDTF without time zone
     *      'revision' (string)
     *      'subject' (string)
     *      'title' (string)
     */
    public function addProperties($values)
    {
        $propsCore = $this->zipPptx->getContent('docProps/core.xml', 'DOMDocument');
        $propsApp = $this->zipPptx->getContent('docProps/app.xml', 'DOMDocument');
        $propsCustom = $this->zipPptx->getContent('docProps/custom.xml', 'DOMDocument');
        $generateCustomRels = false;
        if ($propsCustom === false) {
            $generateCustomRels = true;
            $propsCustom = $this->xmlUtilities->generateDomDocument(OOXMLResources::$customProperties);
            // write the new Override node associated to the new custon.xml file en [Content_Types].xml
            $this->generateOverride('/docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml');
            $this->zipPptx->addContent('docProps/custom.xml', $propsCustom->saveXML());
        }
        $relsRels = $this->zipPptx->getContent('_rels/.rels', 'DOMDocument');

        $prop = new CreateProperties();
        if (!empty($values['title']) || !empty($values['subject']) || !empty($values['creator']) || !empty($values['keywords']) || !empty($values['description']) || !empty($values['category']) || !empty($values['contentStatus']) || !empty($values['created']) || !empty($values['modified']) || !empty($values['lastModifiedBy']) || !empty($values['revision']) ) {
            $propsCore = $prop->createElementProperties($values, $propsCore);
        }
        if (isset($values['contentStatus']) && $values['contentStatus'] == 'Final') {
            $propsCustom = $prop->createPropertiesCustom(array('_MarkAsFinal' => array('boolean' => 'true')), $propsCustom);
        }
        if (!empty($values['Manager']) || !empty($values['Company'])) {
            $propsApp = $prop->createPropertiesApp($values, $propsApp);
        }
        if (!empty($values['custom']) && is_array($values['custom'])) {
            $propsCustom = $prop->createPropertiesCustom($values['custom'], $propsCustom);
            // write the new Override node associated to the new custon.xml file en [Content_Types].xml
            $this->generateOverride('/docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml');
        }
        if ($generateCustomRels) {
            $strCustom = '<Relationship Id="rId' . self::uniqueNumberId(999, 9999) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml" />';
            $tempNode = $relsRels->createDocumentFragment();
            $tempNode->appendXML($strCustom);
            $relsRels->documentElement->appendChild($tempNode);
            // refresh contents
            $this->zipPptx->addContent('ppt/_rels/presentation.xml.rels', $this->pptxRelsPresentationDOM->saveXML());
            $this->pptxRelsPresentationDOM = $this->zipPptx->getContent('ppt/_rels/presentation.xml.rels', 'DOMDocument');
        }

        PhppptxLogger::logger('Set properties.', 'info');

        // refresh contents
        $this->zipPptx->addContent('docProps/core.xml', $propsCore->saveXML());
        $this->zipPptx->addContent('docProps/app.xml', $propsApp->saveXML());
        $this->zipPptx->addContent('docProps/custom.xml', $propsCustom->saveXML());
        $this->zipPptx->addContent('_rels/.rels', $relsRels->saveXML());

        // free DOMDocument resources
        $propsCore = null;
        $propsApp = null;
        $propsCustom = null;
        $relsRels = null;

        PhppptxLogger::logger('Adding properties.', 'info');
    }

    /**
     * Adds a new section
     *
     * @access public
     * @param array $options
     *      'allowDuplicateSectionNames' (bool) if true creates a new section if a section with the same name already exists. Default as false
     *      'name' (string) section name. Default as 'New section'
     *      'moveSlidesWithoutSections' (bool) if true move slides without sections to this new section. Default as false
     *      'position' (int) position. 0 is the first position. -1 (last) as default
     */
    public function addSection($options = array())
    {
        // default options
        if (!isset($options['allowDuplicateSectionNames'])) {
            $options['allowDuplicateSectionNames'] = false;
        }
        if (!isset($options['moveSlidesWithoutSections'])) {
            $options['moveSlidesWithoutSections'] = false;
        }
        if (!isset($options['name']) || empty($options['name'])) {
            $options['name'] = 'New section';
        }
        if (!isset($options['position'])) {
            $options['position'] = -1;
        }

        // get the section tags creating the XML structure if needed
        $nodesExtLst = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'extLst');
        if ($nodesExtLst->length == 0) {
            $newExtLstFragment = $this->pptxPresentationDOM->createDocumentFragment();
            $newExtLstFragment->appendXML('<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" />');
            $nodeExtLst = $this->pptxPresentationDOM->documentElement->appendChild($newExtLstFragment);
        } else {
            $nodeExtLst = $nodesExtLst->item(0);
        }
        $nodesExt = $nodeExtLst->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'ext');
        // check if the correct ext tag exists (with p14:sectionLst tag) or create a new one if needed
        $addNewExt = true;
        if ($nodesExt->length > 0) {
            foreach ($nodesExt as $nodeExt) {
                $nodesP14Section = $nodeExt->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sectionLst');
                if ($nodesP14Section->length > 0) {
                    $addNewExt = false;
                }
            }
        }
        if ($addNewExt) {
            $newExtFragment = $this->pptxPresentationDOM->createDocumentFragment();
            $newExtFragment->appendXML('<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" />');
            $nodeExt = $nodeExtLst->appendChild($newExtFragment);
        }

        $nodesSectionLst = $nodeExtLst->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sectionLst');
        if ($nodesSectionLst->length == 0 && isset($nodeExt)) {
            $newSectionLstFragment = $this->pptxPresentationDOM->createDocumentFragment();
            $newSectionLstFragment->appendXML('<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" />');
            $nodeSectionLst = $nodeExt->appendChild($newSectionLstFragment);
        } else {
            $nodeSectionLst = $nodesSectionLst->item(0);
        }

        $addSection = true;

        // check if the section must be added
        $nodesSection = $nodeSectionLst->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'section');
        if (!$options['allowDuplicateSectionNames'] && $nodesSection->length > 0) {
            foreach ($nodesSection as $nodeSection) {
                if ($nodeSection->hasAttribute('name') && $nodeSection->getAttribute('name') == $options['name']) {
                    $addSection = false;
                    break;
                }
            }
        }

        if ($addSection) {
            // create and add the new section
            $newSectionFragment = $this->pptxPresentationDOM->createDocumentFragment();
            $guid = PhppptxUtilities::generateGUID();
            $newSectionFragment->appendXML('<p14:section id="' . $guid['guid'] . '" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" name="' . $options['name'] . '"><p14:sldIdLst/></p14:section>');
            // add the section in the specific position
            if ($options['position'] < 0 || $options['position'] >= $nodesSection->length) {
                // add the section in the last position or in a position higher than the number of sections add it in the last position
                $nodeSection = $nodeSectionLst->appendChild($newSectionFragment);
            } else {
                // add the slide in a specific position
                $nodesSection->item($options['position'])->parentNode->insertBefore($newSectionFragment, $nodesSection->item($options['position']));
            }
        }

        // refresh contents
        $this->zipPptx->addContent('ppt/presentation.xml', $this->pptxPresentationDOM->saveXML());
        $this->pptxPresentationDOM = $this->zipPptx->getContent('ppt/presentation.xml', 'DOMDocument');

        PhppptxLogger::logger('Add section.', 'info');
    }

    /**
     * Adds a shape
     *
     * @param string $type
     *      accentBorderCallout1, accentBorderCallout2, accentBorderCallout3, accentCallout1, accentCallout2, accentCallout3, actionButtonBackPrevious, actionButtonBeginning, actionButtonBlank, actionButtonDocument, actionButtonEnd, actionButtonForwardNext, actionButtonHelp, actionButtonHome, actionButtonInformation, actionButtonMovie, actionButtonReturn, actionButtonSound, arc
     *      bentArrow, bentConnector2, bentConnector3, bentConnector4, bentConnector5, bentUpArrow, bevel, blockArc, borderCallout1, borderCallout2, borderCallout3, bracePair, bracketPair
     *      callout1, callout2, callout3, can, chartPlus, chartStar, chartX, chevron, chord, circularArrow, cloud, cloudCallout, corner, cornerTabs, cube, curvedConnector2, curvedConnector3, curvedConnector4, curvedConnector5, curvedDownArrow, curvedLeftArrow, curvedRightArrow, curvedUpArrow
     *      decagon, diagStripe, diamond, dodecagon, donut, doubleWave, downArrow, downArrowCallout
     *      ellipse, ellipseRibbon, ellipseRibbon2
     *      flowChartAlternateProcess, flowChartCollate, flowChartConnector, flowChartDecision, flowChartDelay, flowChartDisplay, flowChartDocument, flowChartExtract, flowChartInputOutput, flowChartInternalStorage, flowChartMagneticDisk, flowChartMagneticDrum, flowChartMagneticTape, flowChartManualInput, flowChartManualOperation, flowChartMerge, flowChartMultidocument, flowChartOfflineStorage, flowChartOffpageConnector, flowChartOnlineStorage, flowChartOr, flowChartPredefinedProcess, flowChartPreparation, flowChartProcess, flowChartPunchedCard, flowChartPunchedTape, flowChartSort, flowChartSummingJunction, flowChartTerminator, folderCorner, frame, funnel
     *      gear6, gear9
     *      halfFrame, heart, heptagon, hexagon, homePlate, horizontalScroll
     *      irregularSeal1, irregularSeal2
     *      leftArrow, leftArrowCallout, leftBrace, leftBracket, leftCircularArrow, leftRightArrow, leftRightArrowCallout, leftRightCircularArrow, leftRightRibbon, leftRightUpArrow, leftUpArrow, lightningBolt, line, lineInv
     *      mathDivide, mathEqual, mathMinus, mathMultiply, mathNotEqual, mathPlus, moon
     *      nonIsoscelesTrapezoid, noSmoking, notchedRightArrow
     *      octagon
     *      parallelogram, pentagon, pie, pieWedge, plaque, plaqueTabs, plus
     *      quadArrow, quadArrowCallout
     *      rect, ribbon, ribbon2, rightArrow, rightArrowCallout, rightBrace, rightBracket, round1Rect, round2DiagRect, round2SameRect, roundRect, rtTriangle
     *      smileyFace, snip1Rect, snip2DiagRect, snip2SameRect, snipRoundRect, squareTabs, star10, star12, star16, star24, star32, star4, star5, star6, star7, star8, straightConnector1, stripedRightArrow, sun, swooshArrow
     *      teardrop, trapezoid, triangle
     *      upArrow, upArrowCallout, upDownArrow, upDownArrowCallout, uturnArrow
     *      verticalScroll
     *      wave, wedgeEllipseCallout, wedgeRectCallout, wedgeRoundRectCallout
     * @param array $position
     *      'new' (array) a new shape is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $options
     *      'customGeom' (string) custom geometry
     *      'fillColor' (string) FF0000, 00FFFF,...
     *      'flipH' (bool) flipped horizontally. Default as false
     *      'flipV' (bool) flipped vertically. Default as false
     *      'imageContent' (mixed) image path, base64, stream or GdImage. Image formats: png, jpg, jpeg, gif, bmp, webp
     *      'name' (string) set a name value
     *      'outlineColor' (string) FF0000, 00FFFF,...
     *      'rotation' (int) 60.000ths of a degree
     *      'shapeGuide' (array)
     *          'fmla' (string) shape guide formula
     *          'guide' (string) shape guide name
     *      'tailEnd' (string) arrow, diamond, none, oval, stealth, triangle
     *      'textContents' (PptxFragment)
     *      'textDirection' (string) horz, vert, vert270, wordArtVert, eaVert, mongolianVert, wordArtVertRtl
     *      'verticalAlign' (string) top, middle, bottom
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     * @throws Exception mime option is not set and getimagesizefromstring is not available
     * @throws Exception position not valid
     */
    public function addShape($type, $position, $options = array())
    {
        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
        $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');

        // internal shape ID
        $shapeId = $this->generateUniqueId();
        $options['rId'] = $shapeId;

        $imageId = '';
        if (isset($options['imageContent'])) {
            // internal image ID
            $imageId = $this->generateUniqueId();
            $options['rIdImage'] = $imageId;
        }

        // create and add the new shape
        $shapeElement = new CreateShape();
        $shapeElement->addElementShape($slideDOM, $type, $position, $options);

        // handle external relationships such as hyperlinks
        $externalRelationships = $shapeElement->getExternalRelationships();
        if (count($externalRelationships) > 0) {
            $this->addExternalRelationships($externalRelationships, $activeSlideContent['path'], $slideRelsDOM);
        }

        if (isset($options['imageContent'])) {
            // get image information
            $imageInformation = new ImageUtilities();
            $imageContents = $imageInformation->returnImageContents($options['imageContent'], $options);
            $imageStyles['contents'] = $imageContents;

            // generate the new relationship
            $newRelationship = '<Relationship Id="rId'.$imageId.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img'.$imageId.'.'.$imageContents['extension'].'" />';
            // generate content type if it does not exist yet
            $this->generateDefault($imageContents['extension'], 'image/' . $imageContents['extension']);

            // copy the image with the new name
            $this->zipPptx->addContent('ppt/media/img'.$imageId.'.'.$imageContents['extension'], $imageContents['content']);

            // add the new relationship
            $this->generateRelationship($slideRelsDOM, $newRelationship);
        }

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());
        $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;
        $slideRelsDOM = null;

        PhppptxLogger::logger('Add shape.', 'info');
    }

    /**
     * Adds a new slide
     *
     * @access public
     * @param array $options
     *      'active' (bool) if true set the new slide as the internal active slide. Default as false
     *      'cleanLayoutParagraphContents' (bool) if false do not remove paragraph contents from the layout. Default as true
     *      'cleanSlidePlaceholderTypes' (array) placeholder types to be cleaned from the layout. Default as dt, ftr, hdr, sldNum
     *          Available types: title (Title), body (Body), ctrTitle (Centered Title), subTitle (Subtitle), dt (Date and Time), sldNum (Slide Number), ftr (Footer), hdr (Header), obj (Object), chart (Chart), tbl (Table), clipArt (Clip Art), dgm (Diagram), media (Media), sldImg (Slide Image), pic (Picture)
     *      'layout' (string) layout (name) to be used. If not set, the layout from the active slide is used
     *      'position' (int) slide position. 0 is the first slide. -1 (last) as default. If section is set, use this position to add the slide in the section
     *      'section' (int) section to add the slide. 0 is the first section. -1 is the last section. If not set, the slide is not added in a section
     * @throws Exception layout name doesn't exist
     */
    public function addSlide($options = array())
    {
        // default options
        if (!isset($options['active'])) {
            $options['active'] = false;
        }
        if (!isset($options['cleanLayoutParagraphContents'])) {
            $options['cleanLayoutParagraphContents'] = true;
        }
        if (!isset($options['cleanSlidePlaceholderTypes']) || !is_array($options['cleanSlidePlaceholderTypes'])) {
            $options['cleanSlidePlaceholderTypes'] = array('dt', 'ftr', 'hdr', 'sldNum');
        }
        if (!isset($options['layout'])) {
            // get layout name from the active slide
            $slideContents = $this->zipPptx->getSlides();

            if (count($slideContents) > 0) {
                // at least one slide exists
                $activeSlideContent = $slideContents[$this->activeSlide['position']];
                $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
                $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
                $contentSlideRelsXPath = new DOMXPath($slideRelsDOM);
                $contentSlideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                $nodesSlideLayout = $contentSlideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"]');
                if ($nodesSlideLayout->length > 0) {
                    $targetSlideLayout = '';
                    if ($nodesSlideLayout->item(0)->hasAttribute('Target')) {
                        $targetSlideLayout = $nodesSlideLayout->item(0)->getAttribute('Target');
                    }
                    $nameSlideLayout = '';
                    if (!empty($targetSlideLayout)) {
                        $slideLayoutFilePath = str_replace('../', 'ppt/', $targetSlideLayout);
                        $contentSlideLayout = $this->zipPptx->getContent($slideLayoutFilePath);
                        if (!empty($contentSlideLayout)) {
                            $contentDOMSlideLayout = $this->xmlUtilities->generateDomDocument($contentSlideLayout);
                            $nodesClSd = $contentDOMSlideLayout->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cSld');
                            if ($nodesClSd->length > 0) {
                                if ($nodesClSd->item(0)->hasAttribute('name')) {
                                    $options['layout'] = $nodesClSd->item(0)->getAttribute('name');
                                }
                            }

                            // free DOMDocument resources
                            $contentDOMSlideLayout = null;
                        }
                    }
                }

                // free DOMDocument resources
                $slideRelsDOM = null;
            } else {
                // no slide found. Use the first presentation layout
                $layoutsContents = $this->zipPptx->getLayouts();
                if (count($layoutsContents) > 0) {
                    $options['layout'] = $layoutsContents[0]['name'];
                }
            }
        }
        if (!isset($options['position'])) {
            $options['position'] = -1;
        }

        // check if the layout name exists
        $foundLayout = false;
        $layoutsContents = $this->zipPptx->getLayouts();
        $newLayoutContent = '';
        $newLayoutPath = '';
        foreach ($layoutsContents as $layoutContents) {
            if ($layoutContents['name'] == $options['layout']) {
                $foundLayout = true;

                $newLayoutContent = $layoutContents['content'];
                $newLayoutPath = $layoutContents['path'];
            }
        }

        if (!$foundLayout) {
            PhppptxLogger::logger('The chosen layout name \'' . $options['layout'] . '\' doesn\'t exist. Choose a valid layout.', 'fatal');
        }

        // create and add the new slide

        // get the XML from the layout cleaning the contents to display an empty layout
        $newLayoutContent = $this->xmlUtilities->cleanLayout($newLayoutContent, $options);

        // generate a new ID for the slide
        $newIdRels = $this->generateUniqueId();

        // add Relationship
        $newRelationship = '<Relationship Id="rId'.$newIdRels.'" Target="slides/slide'.$newIdRels.'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>';
        $this->generateRelationship($this->pptxRelsPresentationDOM, $newRelationship);

        // add the slide to the presentation
        $sldIdLstTags = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldIdLst');
        if ($sldIdLstTags->length > 0) {
            // the presentation includes slides
            $newIdSlide = $this->getMaxSlideId();
            $newIdSlide++;
        } else {
            // the presentation doesn't include slides, generate the needed XML content in the correct order
            $newSldIdLstFragment = $this->pptxPresentationDOM->createDocumentFragment();
            $newSldIdLstFragment->appendXML('<p:sldIdLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>');

            // MS PowerPoint default value
            $newIdSlide = 256;

            $notesMasterIdLstTags = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'notesMasterIdLst');
            if ($notesMasterIdLstTags->length > 0) {
                $notesMasterIdLstTags->item(0)->parentNode->insertBefore($newSldIdLstFragment, $notesMasterIdLstTags->item(0)->nextSibling);
            } else {
                $sldMasterIdLstTags = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldMasterIdLst');
                if ($sldMasterIdLstTags->length > 0) {
                    $sldMasterIdLstTags->item(0)->parentNode->insertBefore($newSldIdLstFragment, $sldMasterIdLstTags->item(0)->nextSibling);
                }
            }

            $sldIdLstTags = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldIdLst');
        }

        // add the new slide into the presentation
        $newSlideXML = '<p:sldId id="'.$newIdSlide.'" r:id="rId'.$newIdRels.'" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>';
        $newSlideFragment = $this->pptxPresentationDOM->createDocumentFragment();
        $newSlideFragment->appendXML($newSlideXML);

        $sldIdTags = $sldIdLstTags->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldId');
        if (!isset($options['section'])) {
            if ($options['position'] < 0 || $options['position'] >= $sldIdTags->length) {
                // add the slide in the last position: position lower than 0 or the position is higher than the number of slides
                $sldIdLstTags->item(0)->appendChild($newSlideFragment);
            } else {
                // add the slide in a specific position
                $sldIdTags->item($options['position'])->parentNode->insertBefore($newSlideFragment, $sldIdTags->item($options['position']));
            }
        } else if (isset($options['section'])) {
            // add the slide in the section

            // check if the presentation includes sections and the position is correct
            $sectionLstTags = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sectionLst');
            if ($sectionLstTags->length > 0) {
                $sectionTags = $sectionLstTags->item(0)->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'section');

                $newSldIdFragment = $this->pptxPresentationDOM->createDocumentFragment();
                $newSldIdFragment->appendXML('<p14:sldId id="'.$newIdSlide.'" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" />');

                // get the section tag where to add the new slide
                if ($options['section'] < 0 || $options['section'] >= $sectionTags->length) {
                    // add the slide into the last section
                    $sectionSldIdLstTag = $sectionTags->item($sectionTags->length - 1)->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sldIdLst');
                } else {
                    // add the slide into a speficic section
                    $sectionSldIdLstTag = $sectionTags->item($options['section'])->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sldIdLst');
                }

                $previousSlideId = null;
                if ($sectionSldIdLstTag->length > 0) {
                    $sectionPositionNumber = null;
                    // get previous slide Id and insert the new slide Id in the section tag
                    $sectionIdLstTags = $sectionSldIdLstTag->item(0)->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sldId');
                    if ($options['position'] < 0 || $options['position'] >= $sectionIdLstTags->length) {
                        // get the previous slide Id
                        if ($sectionIdLstTags->length > 0) {
                            $previousSlideId = $sectionIdLstTags->item($sectionIdLstTags->length - 1)->getAttribute('id');
                        }
                        // add the slide in the last position: position lower than 0 or the position is higher than the number of slides
                        $sectionSldIdLstTag->item(0)->appendChild($newSldIdFragment);

                        $sectionPositionNumber = 0;
                    } else {
                        // get the previous slide Id
                        $previousSlideId = $sectionIdLstTags->item($options['position'])->getAttribute('id');

                        // add the slide in a specific position
                        $sectionIdLstTags->item($options['position'])->parentNode->insertBefore($newSldIdFragment, $sectionIdLstTags->item($options['position']));

                        $sectionPositionNumber = $options['position'];
                    }

                    // add the slide in the correct order in the sldIdLst tag, needed by sections

                    // get the p14:sldId tags in p14:sldIdLst
                    $idSlideAdded = false;
                    if ($sectionIdLstTags->length > 0) {
                        // there's at least one sldId tag in the section
                        if ($options['position'] < 0 || $options['position'] >= $sectionIdLstTags->length) {
                            // last slide in the section
                            $idSlideSection = $sectionIdLstTags->item($sectionIdLstTags->length - 1)->getAttribute('id');
                        } else {
                            // specific slide in the section
                            $idSlideSection = $sectionIdLstTags->item($options['position'])->getAttribute('id');
                        }
                        // add the new slide before the slide of the section Id in p:sldIdLst
                        foreach ($sldIdTags as $sldIdTag) {
                            if ($sldIdTag->hasAttribute('id') && $sldIdTag->getAttribute('id') == $previousSlideId) {
                                if ($options['position'] < 0 || $options['position'] >= $sectionIdLstTags->length - 1) {
                                    $sldIdTag->parentNode->insertBefore($newSlideFragment, $sldIdTag->nextSibling);
                                } else {
                                    $sldIdTag->parentNode->insertBefore($newSlideFragment, $sldIdTag);
                                }
                                $idSlideAdded = true;
                                break;
                            }
                        }
                    }

                    if (!$idSlideAdded) {
                        // the slide has not beed added. Get where to add it

                        // get current slide ID
                        $currentSlideId = null;
                        for ($iNextSection = $sectionPositionNumber; $iNextSection < $sectionSldIdLstTag->length; $iNextSection++) {
                            $sectionIdLstTags = $sectionSldIdLstTag->item($iNextSection)->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'sldId');
                            if ($sectionIdLstTags->length > 0) {
                                foreach ($sectionIdLstTags as $sectionIdLstTag) {
                                    if ($sectionIdLstTag->hasAttribute('id') && $sectionIdLstTag->getAttribute('id') != $newIdSlide) {
                                        $currentSlideId = $sectionIdLstTag->getAttribute('id');
                                    }
                                }
                            }
                        }

                        if (!empty($currentSlideId)) {
                            // add the new slide after the last slide of the section Id in p:sldIdLst
                            foreach ($sldIdTags as $sldIdTag) {
                                if ($sldIdTag->hasAttribute('id') && $sldIdTag->getAttribute('id') == $currentSlideId) {
                                    $sldIdTag->parentNode->insertBefore($newSlideFragment, $sldIdTag->nextSibling);
                                }
                            }
                        } else {
                            // there's no other slide ID, append the new slide ID
                            $sldIdLstTags->item(0)->appendChild($newSlideFragment);
                        }
                    }
                }
            }
        }

        // add Override
        $this->generateOverride('/ppt/slides/slide'.$newIdRels.'.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');

        // add the new slide content
        $this->zipPptx->addContent('ppt/slides/slide'.$newIdRels.'.xml', $newLayoutContent);

        // add the new rels content
        $newLayoutPathRels = str_replace('ppt/', '../', $newLayoutPath);
        // check if the layout rels includes image relationships and include them in the new rels
        $extraRelationships = '';
        $layoutRelsPath = str_replace('ppt/slideLayouts/', 'ppt/slideLayouts/_rels/', $newLayoutPath) . '.rels';
        $layoutRelsPathDOM = $this->zipPptx->getContent($layoutRelsPath, 'DOMDocument');
        if ($layoutRelsPathDOM) {
            $layoutRelsPathXPath = new DOMXPath($layoutRelsPathDOM);
            $layoutRelsPathXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
            $relNodesSlideLayout = $layoutRelsPathXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]');
            if ($relNodesSlideLayout->length > 0) {
                foreach ($relNodesSlideLayout as $relNodeSlideLayout) {
                    $extraRelationships .= $relNodeSlideLayout->ownerDocument->saveXML($relNodeSlideLayout);
                }
            }
        }
        $newSlideRelsXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="'.$newLayoutPathRels.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"/>'.$extraRelationships.'</Relationships>';
        $this->zipPptx->addContent('ppt/slides/_rels/slide'.$newIdRels.'.xml.rels', $newSlideRelsXML);

        // refresh contents
        $this->zipPptx->addContent('ppt/presentation.xml', $this->pptxPresentationDOM->saveXML());
        $this->zipPptx->addContent('ppt/_rels/presentation.xml.rels', $this->pptxRelsPresentationDOM->saveXML());
        $this->pptxPresentationDOM = $this->zipPptx->getContent('ppt/presentation.xml', 'DOMDocument');
        $this->pptxRelsPresentationDOM = $this->zipPptx->getContent('ppt/_rels/presentation.xml.rels', 'DOMDocument');

        if (isset($options['active']) && $options['active']) {
            // update the active slide
            $slidesContents = $this->zipPptx->getSlides();

            if ($options['position'] < 0 || $options['position'] >= $sldIdTags->length) {
                // set the internal active slide in the last position
                $this->activeSlide['position'] = count($slidesContents) - 1;
            } else {
                // set the internal active slide in a specific position
                $this->activeSlide['position'] = $options['position'];
            }
        }

        // free DOMDocument resources
        $newLayoutContentDOM = null;

        PhppptxLogger::logger('Add slide.', 'info');
    }

    /**
     * Adds an SVG content
     *
     * @param string $svg SVG path or svg content
     * @param array $position
     *      'new' (array) a new shape is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $svgStyles
     *      'border' (array)
     *          'dash' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *          'color' (string) default as 000000
     *          'width' (int) EMUs (English Metric Unit). Default as 12700 (1pt)
     *      'descr' (string) set a descr value
     *      'rotation' (int)
     * @param array $options
     * @throws Exception ImageMagick extension is not enabled
     * @throws Exception position not valid
     */
    public function addSvg($svg, $position, $svgStyles = array(), $options = array())
    {
        if (!extension_loaded('imagick')) {
            throw new Exception('Install and enable ImageMagick for PHP (https://www.php.net/manual/en/book.imagick.php) to add SVG contents.');
        }

        if (strstr($svg, '<svg')) {
            // SVG is a string content
            $svgContent = $svg;
        } else {
            // SVG is not a string, so it's a file or URL
            $svgContent = file_get_contents($svg);
        }

        // SVG tag
        if (!strstr($svgContent, '<?xml ')) {
            // add an XML tag before the SVG content
            $svgContent = '<?xml version="1.0" encoding="UTF-8" standalone="no"?>' . $svgContent;
        }

        // transform the SVG to PNG using ImageMagick
        $im = new Imagick();
        if (isset($options['resolution']) && isset($options['resolution']['x']) && isset($options['resolution']['y'])) {
            $im->setResolution($options['resolution']['x'], $options['resolution']['y']);
        }
        $im->setBackgroundColor(new ImagickPixel('transparent'));
        $im->readImageBlob($svgContent);
        $im->setImageFormat('png');

        // get image information
        $svgStyles['contents'] = array(
            'content' => $im->getImageBlob(),
            'extension' => 'png',
            'mime' => 'image/png',
            'width' => $im->getImageWidth(),
            'height' => $im->getImageHeight(),
        );

        // make sure that there exists the corresponding content types
        $this->generateDefault('svg', 'svg+xml');
        $this->generateDefault('png', 'image/png');

        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
        $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');

        // internal image ID for the SVG
        $svgId = $this->generateUniqueId();
        $svgStyles['rIdSvg'] = $svgId;

        // internal image ID for the alt image
        $altId = $this->generateUniqueId();
        $svgStyles['rIdAlt'] = $altId;

        // create and add the new pic
        $svgElement = new CreatePic();
        $svgElement->addElementSvg($slideDOM, $position, $svgStyles, $options);

        // generate and add the new relationships
        $newRelationship = '<Relationship Id="rId'.$svgId.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img'.$svgId.'.svg" />';
        $this->generateRelationship($slideRelsDOM, $newRelationship);
        $newRelationship = '<Relationship Id="rId'.$altId.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img'.$altId.'.'.$svgStyles['contents']['extension'].'" />';
        $this->generateRelationship($slideRelsDOM, $newRelationship);

        // copy the contents
        $this->zipPptx->addContent('ppt/media/img'.$svgId.'.svg', $svgContent);
        $this->zipPptx->addContent('ppt/media/img'.$altId.'.'.$svgStyles['contents']['extension'], $svgStyles['contents']['content']);

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());
        $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;
        $slideRelsDOM = null;

        PhppptxLogger::logger('Add SVG.', 'info');
    }

    /**
     * Adds a table
     *
     * @access public
     * @param array $contents array of contents or PptxFragments
     *      'align' (string) left, center, right, justify, distributed
     *      'backgroundColor' (string) HEX color
     *      'border' (array) 'top', 'right', 'bottom', 'left' keys can be used to set borders
     *          'dash' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *          'color' (string) default as 000000. none to avoid adding the color
     *          'width' (int) EMUs (English Metric Unit). Default as 12700 (1pt)
     *      'cellMargin' (array)
     *          'top' (int)
     *          'right' (int)
     *          'bottom' (int)
     *          'left' (int)
     *      'colspan' (int)
     *      'image' (mixed) image path, base64, stream or GdImage. Image formats: png, jpg, jpeg, gif, bmp, webp
     *      'rowspan' (int)
     *      'text' (string|array|PptxFragment) @see addText
     *      'textDirection' (string) horz, vert, vert270, wordArtVert, eaVert, mongolianVert, wordArtVertRtl
     *      'verticalAlign' (string) top, middle, bottom, topCentered, middleCentered, bottomCentered
     *      'wrap' (string) square, none
     * @param array $position
     *      'new' (array) a new position is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $tableStyles
     *      'backgroundColor' (string) HEX color
     *      'bandedColumns' (bool) default as false
     *      'bandedRows' (bool) default as false
     *      'border' (array)
     *          'dash' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *          'color' (string) default as 000000
     *          'width' (int) EMUs (English Metric Unit). Default as 12700 (1pt)
     *      'cellMargin' (array)
     *          'top' (int)
     *          'right' (int)
     *          'bottom' (int)
     *          'left' (int)
     *      'columnWidths' (int|array) column width fix (int) or column width variable (array). If not set, get from the shape size. EMUs (English Metric Unit)
     *      'descr' (string) alt text (descr) value
     *      'firstColumn' (bool) default as false
     *      'headerRow' (bool) default as false
     *      'lastColumn' (bool) default as false
     *      'rtl' (bool) default as false
     *      'style' (string) table style name. Using a template, if the table style doesn't exist, try to import it from the base template
     *      'totalRow' (bool) default as false
     * @param array $rowStyles
     *      'height' (int) EMUs (English Metric Unit)
     * @param array $options
     * @throws Exception position not valid
     * @throws Exception table style doesn't exist
     */
    public function addTable($contents, $position, $tableStyles = array(), $rowStyles = array(), $options = array())
    {
        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
        $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');

        // handle rtl option
        if (self::$rtl && !isset($tableStyles['rtl'])) {
            $tableStyles['rtl'] = true;
        }

        if (isset($tableStyles['style'])) {
            // check if the table style exists

            // get the table styles
            $tableStylesContents = $this->zipPptx->getContentByType('tableStyles');
            if (count($tableStylesContents) == 0) {
                // generate table styles
                $this->zipPptx->addContent('ppt/tableStyles.xml', OOXMLResources::$skeletonTableStyles);
                $this->generateOverride('/ppt/tableStyles.xml', 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml');
                $tableStylesContents = $this->zipPptx->getContentByType('tableStyles');
            }
            $tableStylesDOM = $this->xmlUtilities->generateDomDocument($tableStylesContents[0]['content']);
            $tableStylesXPath = new DOMXPath($tableStylesDOM);
            $tableStylesXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

            $tableStylesNodes = $tableStylesXPath->query('//a:tblStyle[@styleName="'.$tableStyles['style'].'"]');
            if ($tableStylesNodes->length > 0 && $tableStylesNodes->item(0)->hasAttribute('styleId')) {
                // keep the ID to be added when generating the table
                $tableStyles['styleId'] = $tableStylesNodes->item(0)->getAttribute('styleId');
            } else {
                // the style doesn't exist. Check if it can be imported from the base template. Otherwise thrown an Exception
                $baseTemplateStructure = new PptxStructureTemplate();
                $baseTemplateStructureContents = $baseTemplateStructure->getStructure($options);
                $baseTableStylesContents = $baseTemplateStructureContents->getContent('ppt/tableStyles.xml');
                $baseTableStylesDOM = $this->xmlUtilities->generateDomDocument($baseTableStylesContents);
                $baseTableStylesXPath = new DOMXPath($baseTableStylesDOM);
                $baseTableStylesXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
                $baseTableStylesNodes = $baseTableStylesXPath->query('//a:tblStyle[@styleName="'.$tableStyles['style'].'"]');
                if ($baseTableStylesNodes->length > 0) {
                    $newTableStyle = $baseTableStylesNodes->item(0)->ownerDocument->saveXML($baseTableStylesNodes->item(0));
                    // add the required namespaces
                    $newTableStyle = str_replace('<a:tblStyle ', '<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ', $newTableStyle);
                    $newTableStyleFragment = $tableStylesDOM->createDocumentFragment();
                    $newTableStyleFragment->appendXML($newTableStyle);
                    $newTableStyleNode = $tableStylesDOM->documentElement->appendChild($newTableStyleFragment);

                    // refresh contents
                    $this->zipPptx->addContent('ppt/tableStyles.xml', $tableStylesDOM->saveXML());

                    // keep the ID to be added when generating the table
                    $tableStyles['styleId'] = $newTableStyleNode->getAttribute('styleId');
                } else {
                    PhppptxLogger::logger('The style name doesn\'t exist.', 'fatal');
                }

                // free DOMDocument resources
                $baseTableStylesDOM = null;
            }

            // free DOMDocument resources
            $tableStylesDOM = null;
        }

        // parse and add images. Images added in tables do not use the same XML as images added in slides
        foreach ($contents as &$row) {
            foreach ($row as &$cell) {
                if (is_array($cell) && isset($cell['image'])) {
                    // get image information
                    $imageInformation = new ImageUtilities();
                    $imageContents = $imageInformation->returnImageContents($cell['image']);
                    $imageStyles['contents'] = $imageContents;
                    $idImage = $this->generateUniqueId();
                    $imageStyles['rId'] = $idImage;

                    // generate the new relationship
                    $newRelationship = '<Relationship Id="rId'.$idImage.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img'.$idImage.'.'.$imageContents['extension'].'" />';
                    // generate content type if it does not exist yet
                    $this->generateDefault($imageContents['extension'], 'image/' . $imageContents['extension']);

                    // copy the image with the new name
                    $this->zipPptx->addContent('ppt/media/img'.$idImage.'.'.$imageContents['extension'], $imageContents['content']);

                    // add the new relationship
                    $this->generateRelationship($slideRelsDOM, $newRelationship);

                    $cell['image'] = array('xml' => '<a:blipFill><a:blip r:embed="rId'.$idImage.'"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>');
                    $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());
                }
            }
        }

        // create and add the new graphic frame
        $graphicFrameElement = new CreateGraphicFrame();
        $graphicFrameElement->addElementGraphicFrameTable($slideDOM, $contents, $position, $tableStyles, $rowStyles, $options);

        // handle external relationships such as hyperlinks
        $externalRelationships = $graphicFrameElement->getExternalRelationships();
        if (count($externalRelationships) > 0) {
            $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
            $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
            $this->addExternalRelationships($externalRelationships, $activeSlideContent['path'], $slideRelsDOM);

            // refresh contents
            $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

            // free DOMDocument resources
            $slideRelsDOM = null;
        }

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;

        PhppptxLogger::logger('Add table.', 'info');
    }

    /**
     * Adds a text content
     *
     * @access public
     * @param string|array $contents
     *      'text' (string)
     *      'bold' (bool)
     *      'characterSpacing' (int)
     *      'color' (string) HEX color
     *      'font' (string) Arial, Times New Roman ...
     *      'fontSize' (int) 8, 9, 10, 11 ...
     *      'highlight' (string) HEX color
     *      'hyperlink' (string) hyperlink. External, bookmark (#firstslide, #lastslide, #nextslide, #previousslide) or slide (#slide with position)
     *      'italic' (bool)
     *      'lang' (string)
     *      'strikethrough' (bool)
     *      'underline' (string) single
     * @param array $position
     *      'placeholder' (array) the content is added in a placeholder of the layout. One of the following options can be used to get the text box
     *          'name' (string) name
     *          'descr' (string) alt text (descr) value
     *          'position' (int) position by order. 0 is the first order position
     *          'type' (string) title (Title), body (Body), ctrTitle (Centered Title), subTitle (Subtitle)
     *      'new' (array) a new position is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     *          'textBoxStyles' (array) @see addTextBox
     * @param array $paragraphStyles
     *      'align' (string) left, center, right, justify, distributed
     *      'indentation' (int) EMUs (English Metric Unit)
     *      'lineSpacing' (int|float) 1, 1.5, 2...
     *      'listLevel' (int)
     *      'listStyles' (array) @see addList
     *      'marginLeft' (int) EMUs (English Metric Unit) (0 >= and <= 51206400)
     *      'marginRight' (int) EMUs (English Metric Unit) (0 >= and <= 51206400)
     *      'noBullet' (bool) no bullet added. Default as true
     *      'parseLineBreaks' (bool) if true parses the line breaks. Default as false
     *      'rtl' (bool) RTL
     *      'spacingAfter' (int) points (0 >= and <= 158400)
     *      'spacingBefore' (int) points (0 >= and <= 158400)
     * @param array $options
     *      'insertMode' (string) append, replace. If a content exists in the text box, handle how new contents are added. Default as append
     * @throws Exception position not valid
     * @throws Exception hyperlink slide position not valid
     */
    public function addText($contents, $position, $paragraphStyles = array(), $options = array())
    {
        // default values
        if (!isset($options['insertMode'])) {
            $options['insertMode'] = 'append';
        }
        if (!isset($paragraphStyles['noBullet'])) {
            $paragraphStyles['noBullet'] = true;
        }
        if (!isset($paragraphStyles['parseLineBreaks'])) {
            $paragraphStyles['parseLineBreaks'] = false;
        }

        // allow string as $contents instead of an array. Transform string to array
        if (!is_array($contents)) {
            $contentsNormalized = array();
            $contentsNormalized['text'] = $contents;
            $contents = $contentsNormalized;
        }
        // if not using a subarray, generate it
        if (isset($contents['text'])) {
            $contents = array($contents);
        }

        if (!($this instanceof PptxFragment)) {
            // get the internal active slide
            $slideContents = $this->zipPptx->getSlides();
            $activeSlideContent = $slideContents[$this->activeSlide['position']];
            $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
            $slideXPath = new DOMXPath($slideDOM);
            $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

            // create the text tags
            $text = new CreateText();
            $newContentText = $text->createElementText($contents, $paragraphStyles);

            // handle external relationships such as hyperlinks
            $externalRelationships = $text->getExternalRelationships();
            if (count($externalRelationships) > 0) {
                $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
                $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
                $this->addExternalRelationships($externalRelationships, $activeSlideContent['path'], $slideRelsDOM);

                // refresh contents
                $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

                // free DOMDocument resources
                $slideRelsDOM = null;
            }

            // get the box to add the new content
            $nodePSp = $this->getPspBox($position, $slideDOM, $slideXPath, $options);

            // insert the new content
            $nodesTxBody = $nodePSp->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'txBody');
            if ($nodesTxBody->length > 0) {
                // clean existing contents
                $this->cleanPspBox($nodesTxBody->item(0), $slideXPath, $options);

                // append the new contents
                $newTextFragment = $slideDOM->createDocumentFragment();
                $newTextFragment->appendXML($newContentText);
                $nodesTxBody->item(0)->appendChild($newTextFragment);
            }

            // refresh contents
            $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());

            // free DOMDocument resources
            $slideDOM = null;

            PhppptxLogger::logger('Add text content.', 'info');
        } else {
            // create the text tags
            $text = new CreateText();
            $this->pptxML .= $text->createElementText($contents, $paragraphStyles);

            // handle external relationships such as hyperlinks
            $externalRelationships = $text->getExternalRelationships();
            foreach ($externalRelationships as $externalRelationship) {
                $this->addExternalRelationshipFragment($externalRelationship);
            }
        }
    }

    /**
     * Adds a text box
     *
     * @access public
     * @param array $position
     *      'coordinateX' (int) EMUs (English Metric Unit)
     *      'coordinateY' (int) EMUs (English Metric Unit)
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     *      'name' (string) internal name. If not set, a random name is generated
     *      'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $textBoxStyles
     *      'autofit' (string) autofit (default), noautofit, shrink
     *      'border' (array)
     *          'cap' (string) rnd, sq, flat
     *          'dash' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *          'color' (string) default as 000000
     *          'transparency' (int) 0 to 100
     *          'width' (int) EMUs (English Metric Unit). Default as 12700 (1pt)
     *      'columns' (array)
     *          'number' (int)
     *          'spacing' (int) EMUs (English Metric Unit)
     *      'descr' (string)
     *      'fill' (array)
     *          'color' (string) FFFF00, CCCCCC...
     *          'image' (string) image
     *          'imageTransparency' (int)
     *          'imageAsTexture' (bool) Default as false
     *      'margin' (array)
     *          'bottom' (int) EMUs (English Metric Unit)
     *          'left' (int) EMUs (English Metric Unit)
     *          'right' (int) EMUs (English Metric Unit)
     *          'top' (int) EMUs (English Metric Unit)
     *      'rotation' (int)
     *      'textDirection' (string) horz, vert, vert270, wordArtVert, eaVert, mongolianVert, wordArtVertRtl
     *      'verticalAlign' (string) top, middle, bottom, topCentered, middleCentered, bottomCentered
     *      'wrap' (string) square, none
     * @param array $options
     * @throws Exception position not valid
     */
    public function addTextBox($position, $textBoxStyles = array(), $options = array())
    {
        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        if (isset($textBoxStyles['fill']) && isset($textBoxStyles['fill']['image'])) {
            // generate an ID for the fill image
            $idImage = $this->generateUniqueId();
            $textBoxStyles['fill']['imageId'] = $idImage;
        }

        // create and add the new text box
        $textBox = new CreateTextBox();
        $textBox->addElementTextBox($slideDOM, $position, $textBoxStyles, $options);

        if (isset($textBoxStyles['fill']) && isset($textBoxStyles['fill']['image'])) {
            // add the fill image
            $imageInformation = new ImageUtilities();
            $imageContents = $imageInformation->returnImageContents($textBoxStyles['fill']['image'], $options);

            // generate the new relationship
            $newRelationship = '<Relationship Id="rId' . $textBoxStyles['fill']['imageId'] . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img' . $textBoxStyles['fill']['imageId'] . '.' . $imageContents['extension'] . '" />';
            // generate content type if it does not exist yet
            $this->generateDefault($imageContents['extension'], 'image/' . $imageContents['extension']);

            // copy the image with the new name
            $this->zipPptx->addContent('ppt/media/img' . $textBoxStyles['fill']['imageId'] . '.' . $imageContents['extension'], $imageContents['content']);

            // add the new relationship
            $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
            $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
            $this->generateRelationship($slideRelsDOM, $newRelationship);

            // refresh contents
            $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

            // free DOMDocument resources
            $slideRelsDOM = null;
        }

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;

        PhppptxLogger::logger('Add text box.', 'info');
    }

    /**
     * Adds a text box connector
     *
     * @param array $position
     *      'coordinateX' (int) EMUs (English Metric Unit). Automatically calculated from the start and end connections if not set
     *      'coordinateY' (int) EMUs (English Metric Unit). Automatically calculated from the start and end connections if not set
     *      'sizeX' (int) EMUs (English Metric Unit). Automatically calculated from the start and end connections if not set
     *      'sizeY' (int) EMUs (English Metric Unit). Automatically calculated from the start and end connections if not set
     *      'name' (string) internal name. If not set, a random name is generated
     *      'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $connection
     *      'start' (int|string) id or internal name
     *      'end' (int|string) id or internal name
     *      'positionStart' (string) connection position: top, left, right (default), bottom. Only used when $position is calculated automatically. If not set, the position is detected automatically
     *      'positionEnd' (string) connection position: top, left (default), right, bottom. Only used when $position is calculated automatically. If not set, the position is detected automatically
     *      'flipH' (bool) flipped horizontally. Default as false
     *      'flipV' (bool) flipped vertically. Default as false
     * @param array $options
     *      'color' (string) FF0000, 00FFFF,...
     *      'dash' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *      'geom' (string) bentConnector2, bentConnector3, bentConnector4, bentConnector5, curvedConnector2, curvedConnector3, curvedConnector4, curvedConnector5, straightConnector1 (default)
     *      'lineWidth' (int) EMUs (English Metric Unit). 12700 = 1pt
     *      'rotation' (int) 60.000ths of a degree
     *      'shapeGuide' (array)
     *          'fmla' (string) shape guide formula
     *          'guide' (string) shape guide name
     *      'tailEnd' (string) arrow, diamond, none, oval, stealth, triangle (default)
     * @throws Exception method not available
     * @throws Exception not valid connections
     * @throws Exception position not valid
     */
    public function addTextBoxConnector($position, $connection, $options = array())
    {
        if (!file_exists(__DIR__ . '/CreateTextBoxConnector.php')) {
            PhppptxLogger::logger('This method is not available for your license.', 'fatal');
        }

        if (!isset($connection['start']) || !isset($connection['end'])) {
            PhppptxLogger::logger('The chosen connections are not valid. Use valid connections.', 'fatal');
        }

        // default options
        if (!isset($options['geom'])) {
            $options['geom'] = 'straightConnector1';
        }
        if (!isset($options['tailEnd'])) {
            $options['tailEnd'] = 'triangle';
        }
        if (!isset($connection['positionStart'])) {
            $connection['positionStart'] = 'right';
        }
        if (!isset($connection['positionStart'])) {
            $connection['positionStart'] = 'right';
        }
        if (!isset($connection['positionEnd'])) {
            $connection['positionEnd'] = 'left';
        }

        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // if internal name is used, get and use the ID number as required
        if (!is_numeric($connection['start'])) {
            $nodesCNvPr = $slideXPath->query('//p:sp//p:cNvPr[@name="'.$connection['start'].'"]');
            if ($nodesCNvPr->length > 0) {
                $connection['start'] = $nodesCNvPr->item(0)->getAttribute('id');
            }
        }
        if (!is_numeric($connection['end'])) {
            $nodesCNvPr = $slideXPath->query('//p:sp//p:cNvPr[@name="'.$connection['end'].'"]');
            if ($nodesCNvPr->length > 0) {
                $connection['end'] = $nodesCNvPr->item(0)->getAttribute('id');
            }
        }

        if (!isset($position['coordinateX']) || !isset($position['coordinateY']) || !isset($position['sizeX']) || !isset($position['sizeY'])) {
            // get position from start and end connections
            $nodesXfrmStart = $slideXPath->query('//p:sp[.//p:nvSpPr//p:cNvPr[@id="'.$connection['start'].'"]]//a:xfrm | //p:sp[.//p:nvSpPr//p:cNvPr[@name="'.$connection['start'].'"]]//a:xfrm');
            $nodesXfrmEnd = $slideXPath->query('//p:sp[.//p:nvSpPr//p:cNvPr[@id="'.$connection['end'].'"]]//a:xfrm | //p:sp[.//p:nvSpPr//p:cNvPr[@name="'.$connection['end'].'"]]//a:xfrm');
            if ($nodesXfrmStart->length > 0 && $nodesXfrmEnd->length > 0) {
                $nodesOffStart = $slideXPath->query('.//a:off', $nodesXfrmStart->item(0));
                $nodesExtStart = $slideXPath->query('.//a:ext', $nodesXfrmStart->item(0));
                $nodesOffEnd = $slideXPath->query('.//a:off', $nodesXfrmEnd->item(0));
                $nodesExtEnd = $slideXPath->query('.//a:ext', $nodesXfrmEnd->item(0));

                if ($nodesOffStart->length > 0 && $nodesExtStart->length > 0 && $nodesOffEnd->length > 0 && $nodesExtEnd->length > 0) {
                    if (!isset($position['coordinateX'])) {
                        if ($connection['positionStart'] == 'right') {
                            $position['coordinateX'] = (int)$nodesOffStart->item(0)->getAttribute('x') + (int)$nodesExtStart->item(0)->getAttribute('cx');
                        } else if ($connection['positionStart'] == 'left') {
                            $position['coordinateX'] = (int)$nodesOffStart->item(0)->getAttribute('x');
                        } else if ($connection['positionStart'] == 'top' || $connection['positionStart'] == 'bottom') {
                            $position['coordinateX'] = (int)$nodesOffStart->item(0)->getAttribute('x') + ((int)$nodesExtStart->item(0)->getAttribute('cx') / 2);
                        }
                    }

                    if (!isset($position['coordinateY'])) {
                        if ($connection['positionStart'] == 'left' || $connection['positionStart'] == 'right') {
                            $position['coordinateY'] = (int)$nodesOffStart->item(0)->getAttribute('y') + ((int)$nodesExtStart->item(0)->getAttribute('cy') / 2);
                        } else if ($connection['positionStart'] == 'top') {
                            $position['coordinateY'] = (int)$nodesOffStart->item(0)->getAttribute('y');
                        } else if ($connection['positionStart'] == 'bottom') {
                            $position['coordinateY'] = (int)$nodesOffStart->item(0)->getAttribute('y') + (int)$nodesExtStart->item(0)->getAttribute('cy');
                        }
                    }

                    if (!isset($position['sizeX'])) {
                        $position['sizeX'] = (int)$nodesOffEnd->item(0)->getAttribute('x') - $position['coordinateX'];
                        if ($connection['positionEnd'] == 'right') {
                            $position['sizeX'] += (int)$nodesExtEnd->item(0)->getAttribute('cx');
                        } else if ($connection['positionEnd'] == 'top' || $connection['positionEnd'] == 'bottom') {
                            $position['sizeX'] += ((int)$nodesExtEnd->item(0)->getAttribute('cx') / 2);
                        }
                        if ($position['sizeX'] < 0) {
                            // negative size, set the correct values based on flip and size
                            $position['sizeX'] = abs($position['sizeX']);
                            $position['coordinateX'] -= $position['sizeX'];
                            if (!isset($connection['flipH'])) {
                                $connection['flipH'] = true;
                            }
                        }
                    }

                    if (!isset($position['sizeY'])) {
                        $position['sizeY'] = (int)$nodesOffEnd->item(0)->getAttribute('y') - (int)$nodesOffStart->item(0)->getAttribute('y');
                        if ($connection['positionEnd'] == 'top') {
                            $position['sizeY'] -= (int)$nodesExtEnd->item(0)->getAttribute('cy') / 2;
                        } else if ($connection['positionEnd'] == 'bottom') {
                            $position['sizeY'] += (int)$nodesExtEnd->item(0)->getAttribute('cy') / 2;
                        }
                        if ($connection['positionStart'] == 'top') {
                            $position['sizeY'] += (int)$nodesExtEnd->item(0)->getAttribute('cy') / 2;
                        }
                        if ($connection['positionStart'] == 'bottom') {
                            $position['sizeY'] -= (int)$nodesExtEnd->item(0)->getAttribute('cy') / 2;
                        }
                        if ($position['sizeY'] < 0) {
                            // negative size, set the correct values based on flip and size
                            $position['sizeY'] = abs($position['sizeY']);
                            $position['coordinateY'] -= $position['sizeY'];
                            if (!isset($connection['flipV'])) {
                                $connection['flipV'] = true;
                            }
                        }
                    }
                }
            }
        }

        // internal connector ID
        $shapeId = $this->generateUniqueId();
        $options['rId'] = $shapeId;

        // create and add the new connector
        $shapeElementConnector = new CreateTextBoxConnector();
        $shapeElementConnector->addElementTextBoxConnector($slideDOM, $position, $connection, $options);

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;

        PhppptxLogger::logger('Add text box connector.', 'info');
    }

    /**
     * Adds a video
     *
     * @param string $video video path. Video formats: avi, mkv, mp4, wmv
     * @param array $position
     *      'new' (array) a new position is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $videoStyles
     *      'image' (array)
     *          'image' (string) image to be used as preview. Set a default one if not set
     *          'mime' (string) forces a mime (image/jpg, image/jpeg, image/png, image/gif, image/webp)
     *      'mime' (string) forces a mime (video/mp4, video/x-msvideo, video/x-ms-wmv, video/unknown)
     * @param array $options
     *      'descr' (string) alt text (descr) value
     * @throws Exception video doesn't exist
     * @throws Exception video format is not supported
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     * @throws Exception position not valid
     */
    public function addVideo($video, $position, $videoStyles = array(), $options = array())
    {
        // get video information
        $videoInformation = new VideoUtilities();
        $videoContents = $videoInformation->returnVideoContents($video, $options);
        $videoStyles['contents'] = $videoContents;
        $options['type'] = 'video';

        // add the video
        $this->addMedia($videoContents, $position, $videoStyles, $options);

        PhppptxLogger::logger('Add video.', 'info');
    }

    /**
     * Gets active slide information
     *
     * @access public
     * @return array
     */
    public function getActiveSlideInformation()
    {
        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        $slidePlaceholders = array();

        $nodesCNvPr = $slideXPath->query('//p:sp//p:cNvPr');
        foreach ($nodesCNvPr as $nodeCNvPr) {
            $slidePlaceholder = array();
            if ($nodeCNvPr->hasAttribute('name')) {
                $slidePlaceholder['name'] = $nodeCNvPr->getAttribute('name');
            }
            if ($nodeCNvPr->hasAttribute('id')) {
                $slidePlaceholder['id'] = $nodeCNvPr->getAttribute('id');
            }
            if ($nodeCNvPr->hasAttribute('descr')) {
                $slidePlaceholder['descr'] = $nodeCNvPr->getAttribute('descr');
            }

            $slidePlaceholders[] = $slidePlaceholder;
        }

        return array(
            'placeholders' => $slidePlaceholders,
        );
    }

    /**
     * Clones elements using PptxPath queries
     *
     * @access public
     * @param array $referenceNode reference node to clone
     *      'target' (string) slide (default)
     *      'type' (string) audio, image, paragraph, run, shape (text box) (default), slide, table, table-row, video
     *      'contains' (string) for paragraph, run, shape, table, table-row types
     *      'occurrence' (int) exact occurrence (1 is the first occurrence), (array) occurrences, (string) range of contents (e.g.: 2..9, 2.., ..9) or first() or last()
     *      'attributes' (array) node or descendant attributes
     *      'parent' (string) immediate children (default as '/', any parent) or any other parent (a:tbl/, p:sp/...)
     *      'rootParent' (string) root parent. Default as p:spTree for slide elements and p:presentation for presentation elements
     *      'customQuery' (string) if set overwrites all previous references. It must be a valid XPath query
     * @param array $options
     *      'name' (string) custom internal name. If not set, a random name is generated
     * @throws Exception method not available
     */
    public function cloneElement($referenceNode, $options = array())
    {
        if (!file_exists(__DIR__ . '/PptxPath.php')) {
            PhppptxLogger::logger('This method is not available for your license.', 'fatal');
        }

        // default options
        if (!isset($referenceNode['target'])) {
            $referenceNode['target'] = 'slide';
        }
        if (!isset($referenceNode['type'])) {
            $referenceNode['type'] = 'shape';
        }

        $contentsType = $this->getContentsDOM($referenceNode['target'], $referenceNode);

        // get the referenceNode
        $query = $this->getPptxQuery($referenceNode);
        $contentNodes = $contentsType['xpath']->query($query);

        if ($referenceNode['type'] == 'slide') {
            foreach ($contentNodes as $contentNode) {
                // clone slides
                $newIdSlide = $this->getMaxSlideId();
                $newIdSlide++;

                // clone the slide node
                $newNode = $contentNode->cloneNode(true);

                $newNode->setAttribute('id', $newIdSlide);
                $newIdRels = $this->generateUniqueId();
                $newNode->setAttribute('r:id', 'rId' . $newIdRels);

                // append the new slide after the cloned slide
                $contentNode->parentNode->insertBefore($newNode, $contentNode->nextSibling);

                // check if the slide is in a presentation and clone it if needed
                $sectionCloneSldId = $contentsType['xpath']->query('//p14:section/p14:sldIdLst/p14:sldId[@id="'.$contentNode->getAttribute('id').'"]');
                if ($sectionCloneSldId->length > 0) {
                    $newSectionCloneSldId = $sectionCloneSldId->item(0)->cloneNode(true);
                    $newSectionCloneSldId->setAttribute('id', $newIdSlide);
                    $sectionCloneSldId->item(0)->parentNode->insertBefore($newSectionCloneSldId, $sectionCloneSldId->item(0)->nextSibling);
                }

                // clone slide contents

                // slide.xml and slide.xml.rels
                $slideContents = $this->zipPptx->getSlides();
                foreach ($slideContents as $slideContent) {
                    if ($slideContent['id'] == $contentNode->getAttribute('r:id')) {
                        $slideContentClone = $this->zipPptx->getContent($slideContent['path']);
                        if ($slideContentClone) {
                            $this->zipPptx->addContent('ppt/slides/slide' . $newIdRels . '.xml', $slideContentClone);
                        }
                        $slideRelsContentClone = $this->zipPptx->getContent(str_replace('ppt/slides/', 'ppt/slides/_rels/', $slideContent['path']) . '.rels');
                        if ($slideRelsContentClone) {
                            // clone external XML relationships: notes, comments, charts, diagrams
                            $slideRelsContentCloneDOM = $this->xmlUtilities->generateDomDocument($slideRelsContentClone);
                            $slideRelsContentCloneXPath = new DOMXPath($slideRelsContentCloneDOM);
                            $slideRelsContentCloneXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                            $nodesRelationship = $slideRelsContentCloneXPath->query('//xmlns:Relationship');
                            foreach ($nodesRelationship as $nodeRelationship) {
                                if ($nodeRelationship->hasAttribute('Type')) {
                                    $relationshipType = $nodeRelationship->getAttribute('Type');
                                    if ($relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments') {
                                        // comments
                                        $nodeRelationshipContent = $this->zipPptx->getContent('ppt/' . str_replace('../', '', str_replace('ppt/', '', $nodeRelationship->getAttribute('Target'))));
                                        $newIdRelationship = $this->generateUniqueId();
                                        $slideRelsContentClone = str_replace($nodeRelationship->getAttribute('Target'), '../comments/comment' . $newIdRelationship . '.xml', $slideRelsContentClone);
                                        $this->zipPptx->addContent('ppt/comments/comment' . $newIdRelationship . '.xml', $nodeRelationshipContent);
                                        $this->generateOverride('/ppt/comments/comment' . $newIdRelationship . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml');
                                    } elseif ($relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide') {
                                        // notes
                                        $nodeRelationshipContent = $this->zipPptx->getContent('ppt/' . str_replace('../', '', str_replace('ppt/', '', $nodeRelationship->getAttribute('Target'))));
                                        $newIdRelationship = $this->generateUniqueId();
                                        $slideRelsContentClone = str_replace($nodeRelationship->getAttribute('Target'), '../notesSlides/notesSlide' . $newIdRelationship . '.xml', $slideRelsContentClone);
                                        $this->zipPptx->addContent('ppt/notesSlides/notesSlide' . $newIdRelationship . '.xml', $nodeRelationshipContent);
                                        $this->generateOverride('/ppt/notesSlides/notesSlide' . $newIdRelationship . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');

                                        // note rels
                                        $nodeNotesRelsContent = $this->zipPptx->getContent('ppt/' . str_replace('../notesSlides/', 'notesSlides/_rels/', str_replace('ppt/', '', $nodeRelationship->getAttribute('Target'))) . '.rels');
                                        $nodeNotesRelsContent = str_replace($nodesRelationship->item(0)->getAttribute('Target'), 'slides/slide' . $newIdRelationship . '.xml', $nodeNotesRelsContent);
                                        $this->zipPptx->addContent('ppt/notesSlides/_rels/notesSlides' . $newIdRelationship . '.xml.rels', $nodeNotesRelsContent);
                                    } elseif (
                                        $relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout' ||
                                        $relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData' ||
                                        $relationshipType == 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing' ||
                                        $relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors' ||
                                        $relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle'
                                    ) {
                                        // diagrams
                                        $diagramScope = null;
                                        if ($relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout') {
                                            $diagramScope = 'layout';
                                        } elseif ($relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData') {
                                            $diagramScope = 'data';
                                        } elseif ($relationshipType == 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing') {
                                            $diagramScope = 'drawing';
                                        } elseif ($relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors') {
                                            $diagramScope = 'colors';
                                        } elseif ($relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle') {
                                            $diagramScope = 'quickStyle';
                                        }

                                        $nodeRelationshipContent = $this->zipPptx->getContent('ppt/' . str_replace('../', '', str_replace('ppt/', '', $nodeRelationship->getAttribute('Target'))));
                                        $newIdRelationship = $this->generateUniqueId();
                                        $slideRelsContentClone = str_replace($nodeRelationship->getAttribute('Target'), '../diagrams/' . $diagramScope . $newIdRelationship . '.xml', $slideRelsContentClone);
                                        $this->zipPptx->addContent('ppt/diagrams/' . $diagramScope . $newIdRelationship . '.xml', $nodeRelationshipContent);
                                        $this->generateOverride('/ppt/diagrams/' . $diagramScope . $newIdRelationship . '.xml', 'application/vnd.openxmlformats-officedocument.drawingml.diagram'.ucfirst($diagramScope).'+xml');
                                    } elseif ($relationshipType == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart') {
                                        // charts
                                        $nodeRelationshipContent = $this->zipPptx->getContent('ppt/' . str_replace('../', '', str_replace('ppt/', '', $nodeRelationship->getAttribute('Target'))));
                                        $newIdRelationship = $this->generateUniqueId();
                                        $slideRelsContentClone = str_replace($nodeRelationship->getAttribute('Target'), '../charts/chart' . $newIdRelationship . '.xml', $slideRelsContentClone);
                                        $this->zipPptx->addContent('ppt/charts/chart' . $newIdRelationship . '.xml', $nodeRelationshipContent);
                                        if ($nodeRelationship->getAttribute('Type') == 'http://schemas.microsoft.com/office/2014/relationships/chartEx') {
                                            // extended chart
                                            $this->generateOverride('/ppt/charts/chart' . $newIdRelationship . '.xml', 'application/vnd.ms-office.chartex+xml');
                                        } else {
                                            // other chart
                                            $this->generateOverride('/ppt/charts/chart' . $newIdRelationship . '.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
                                        }

                                        // rels
                                        $nodeChartRelsContent = $this->zipPptx->getContent('ppt/' . str_replace('../charts/', 'charts/_rels/', str_replace('ppt/', '', $nodeRelationship->getAttribute('Target'))) . '.rels');
                                        $nodeChartRelsContentDOM = $this->xmlUtilities->generateDomDocument($nodeChartRelsContent);
                                        $nodeChartRelsContentXPath = new DOMXPath($nodeChartRelsContentDOM);
                                        $nodeChartRelsContentXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                                        $nodesRelationshipChart = $nodeChartRelsContentXPath->query('//xmlns:Relationship');
                                        foreach ($nodesRelationshipChart as $nodeRelationshipChart) {
                                            if ($nodeRelationshipChart->getAttribute('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package') {
                                                $chartContent = $this->zipPptx->getContent('ppt/' . str_replace('../embeddings/', 'embeddings/', str_replace('ppt/', '', $nodeRelationshipChart->getAttribute('Target'))));
                                                $this->zipPptx->addContent('ppt/embeddings/Microsoft_Excel_Worksheet' . $newIdRelationship . '.xlsx', $chartContent);

                                                $nodeChartRelsContent = str_replace($nodeRelationshipChart->getAttribute('Target'), '../embeddings/Microsoft_Excel_Worksheet' . $newIdRelationship . '.xlsx', $nodeChartRelsContent);
                                            } else if ($nodeRelationshipChart->getAttribute('Type') == 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle') {
                                                $chartContent = $this->zipPptx->getContent('ppt/charts/' . str_replace('ppt/', '', $nodeRelationshipChart->getAttribute('Target')));
                                                $this->zipPptx->addContent('ppt/charts/colors' . $newIdRelationship . '.xml', $chartContent);

                                                $nodeChartRelsContent = str_replace(str_replace('ppt/', '', $nodeRelationshipChart->getAttribute('Target')), 'colors' . $newIdRelationship . '.xml', $nodeChartRelsContent);
                                                $this->generateOverride('/ppt/charts/colors' . $newIdRelationship . '.xml', 'application/vnd.ms-office.chartcolorstyle+xml');
                                            } else if ($nodeRelationshipChart->getAttribute('Type') == 'http://schemas.microsoft.com/office/2011/relationships/chartStyle') {
                                                $chartContent = $this->zipPptx->getContent('ppt/charts/' . str_replace('ppt/', '', $nodeRelationshipChart->getAttribute('Target')));
                                                $this->zipPptx->addContent('ppt/charts/style' . $newIdRelationship . '.xml', $chartContent);

                                                $nodeChartRelsContent = str_replace(str_replace('ppt/', '', $nodeRelationshipChart->getAttribute('Target')), 'style' . $newIdRelationship . '.xml', $nodeChartRelsContent);
                                                $this->generateOverride('/ppt/charts/style' . $newIdRelationship . '.xml', 'application/vnd.ms-office.chartstyle+xml');
                                            }
                                        }

                                        // free DOMDocument resources
                                        $nodeChartRelsContentDOM = null;

                                        $this->zipPptx->addContent('ppt/charts/_rels/chart' . $newIdRelationship . '.xml.rels', $nodeChartRelsContent);
                                    }
                                }
                            }

                            $this->zipPptx->addContent('ppt/slides/_rels/slide' . $newIdRels . '.xml.rels', $slideRelsContentClone);

                            // free DOMDocument resources
                            $slideRelsContentCloneDOM = null;
                        }

                        break;
                    }
                }

                // slide in presentation.xml.rels
                $newRelationship = '<Relationship Id="rId'.$newIdRels.'" Target="slides/slide'.$newIdRels.'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>';
                $this->generateRelationship($this->pptxRelsPresentationDOM, $newRelationship);
                $this->zipPptx->addContent('ppt/_rels/presentation.xml.rels', $this->pptxRelsPresentationDOM->saveXML());
                $this->pptxRelsPresentationDOM = $this->zipPptx->getContent('ppt/_rels/presentation.xml.rels', 'DOMDocument');

                // ContentType
                $this->generateOverride('/ppt/slides/slide'.$newIdRels.'.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
            }
        } else {
            // get slide cNvPr ids
            $idsCnvPr = array();
            $nodesCnvPr = $contentsType['xpath']->query('//p:cNvPr');
            foreach ($nodesCnvPr as $nodeCnvPr) {
                if ($nodeCnvPr->hasAttribute('id')) {
                    $idsCnvPr[] = $nodeCnvPr->getAttribute('id');
                }
            }

            foreach ($contentNodes as $contentNode) {
                // clone the node
                $newNode = $contentNode->cloneNode(true);

                // generate and set a new cNvPr id and name
                $newPlaceholderId = null;
                while (!isset($newPlaceholderId)) {
                    $randomId = mt_rand(999, 9999999);
                    if (!in_array($randomId, $idsCnvPr)) {
                        $idsCnvPr[] = $randomId;
                        $newPlaceholderId = $randomId;
                    }
                }
                $cNvPrNodes = $newNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cNvPr');
                if ($cNvPrNodes->length > 0) {
                    $cNvPrNodes->item(0)->setAttribute('id', $newPlaceholderId);

                    if ($cNvPrNodes->item(0)->hasAttribute('name')) {
                        $newCNvPrNodeName = $cNvPrNodes->item(0)->getAttribute('name') . $this->generateUniqueId();
                        if (isset($options['name'])) {
                            $newCNvPrNodeName = $options['name'];
                        }
                        $cNvPrNodes->item(0)->setAttribute('name', $newCNvPrNodeName);
                    }
                }

                // append the new node
                $contentNode->parentNode->appendChild($newNode);
            }
        }

        // refresh contents
        $this->zipPptx->addContent($contentsType['path'], $contentsType['dom']->saveXML());
        if (in_array($referenceNode['type'], array('slide'))) {
            $this->pptxPresentationDOM = $this->zipPptx->getContent('ppt/presentation.xml', 'DOMDocument');
        }

        // free DOMDocument resources
        $contentsType['dom'] = null;
    }

    /**
     * Creates a table style
     *
     * @access public
     * @param string $name table style name
     * @param array $styles
     *      'backgroundColor' (string) HEX color
     *      'bold' (bool)
     *      'border' (array) 'top', 'right', 'bottom', 'left', 'insideH' and 'insideV' keys can be used to set borders
     *          'dash' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *          'color' (string) default as 000000. none to avoid adding the color
     *          'width' (int) EMUs (English Metric Unit). Default as 12700 (1pt)
     *      'italic' (bool)
     *
     * Available types to be used in $styles:
     *      'wholeTbl' (array) whole table
     *      'band1H' (array) band 1 horizontal
     *      'band2H' (array) band 2 horizontal
     *      'band1V' (array) band 1 vertical
     *      'band2V' (array) band 2 vertical
     *      'lastCol' (array) last column
     *      'firstCol' (array) first column
     *      'lastRow' (array) last row
     *      'firstRow' (array) first row
     *      'neCell' (array) northeast cell
     *      'seCell' (array) southeast cell
     *      'swCell' (array) southwest cell
     *      'nwCell' (array) northwest cell
     * @throws Exception duplicated table style name
     */
    public function createTableStyle($name, $styles)
    {
        // get the table styles
        $tableStylesContents = $this->zipPptx->getContentByType('tableStyles');
        if (count($tableStylesContents) == 0) {
            // generate table styles
            $this->zipPptx->addContent('ppt/tableStyles.xml', OOXMLResources::$skeletonTableStyles);
            $this->generateOverride('/ppt/tableStyles.xml', 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml');
            $tableStylesContents = $this->zipPptx->getContentByType('tableStyles');
        }
        $tableStylesDOM = $this->xmlUtilities->generateDomDocument($tableStylesContents[0]['content']);
        $tableStylesXPath = new DOMXPath($tableStylesDOM);
        $tableStylesXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

        $tableStylesNodes = $tableStylesXPath->query('//a:tblStyle[@styleName="'.$name.'"]');

        if ($tableStylesNodes->length > 0) {
            PhppptxLogger::logger('The table style name "' . $name . '" already exists.', 'fatal');
        }

        $tableStyle = new CreateTableStyle();
        $tableStyleContent = $tableStyle->createTableStyle($name, $styles);

        // add the new table style
        $newTableStyleFragment = $tableStylesDOM->createDocumentFragment();
        $newTableStyleFragment->appendXML($tableStyleContent);
        $newTableStyleNode = $tableStylesDOM->documentElement->appendChild($newTableStyleFragment);

        // refresh contents
        $this->zipPptx->addContent('ppt/tableStyles.xml', $tableStylesDOM->saveXML());

        // free DOMDocument resources
        $tableStylesDOM = null;

        PhppptxLogger::logger('Create table style.', 'info');
    }

    /**
     * Customizes element styles
     *
     * @access public
     * @param array $referenceNode reference node
     *      'target' (string) slide (default), slideLayout
     *      'type' (string) audio, chart, diagram, image, paragraph, run, shape (text box) (default), table, table-row, table-cell, table-cell-paragraph, video
     *      'contains' (string) for paragraph, run, shape, table, table-row, table-cell, table-cell-paragraph types
     *      'occurrence' (int) exact occurrence (1 is the first occurrence), (array) occurrences, (string) range of contents (e.g.: 2..9, 2.., ..9) or first() or last()
     *      'attributes' (array) node or descendant attributes
     *      'parent' (string) immediate children (default as '/', any parent) or any other parent (a:tbl/, p:sp/...)
     *      'rootParent' (string) root parent. Default as p:spTree for slide elements and p:presentation for presentation elements
     *      'customQuery' (string) if set overwrites all previous references. It must be a valid XPath query
     * @param array $options Style options to apply to the content
     * Options audio:
     *      'coordinateX' (int) EMUs (English Metric Unit)
     *      'coordinateY' (int) EMUs (English Metric Unit)
     *      'name' (string) internal name
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     * Options chart:
     *      'coordinateX' (int) EMUs (English Metric Unit)
     *      'coordinateY' (int) EMUs (English Metric Unit)
     *      'name' (string) internal name
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     * Options diagram:
     *      'coordinateX' (int) EMUs (English Metric Unit)
     *      'coordinateY' (int) EMUs (English Metric Unit)
     *      'name' (string) internal name
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     * Options image:
     *      'coordinateX' (int) EMUs (English Metric Unit)
     *      'coordinateY' (int) EMUs (English Metric Unit)
     *      'name' (string) internal name
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     * Options paragraph:
     *      'align' (string) left, center, right, justify, distributed
     *      'bold' (bool)
     *      'characterSpacing' (int)
     *      'color' (string) HEX color
     *      'font' (string) Arial, Times New Roman ...
     *      'fontSize' (int) 8, 9, 10, 11 ...
     *      'highlight' (string) HEX color
     *      'indentation' (int) EMUs (English Metric Unit)
     *      'italic' (bool)
     *      'lang' (string)
     *      'lineSpacing' (int|float) 1, 1.5, 2...
     *      'marginLeft' (int) EMUs (English Metric Unit) (0 >= and <= 51206400)
     *      'marginRight' (int) EMUs (English Metric Unit) (0 >= and <= 51206400)
     *      'rtl' (bool) RTL
     *      'spacingAfter' (int) points (0 >= and <= 158400)
     *      'spacingBefore' (int) points (0 >= and <= 158400)
     *      'strikethrough' (bool)
     *      'underline' (string) single
     * Options run:
     *      'bold' (bool)
     *      'characterSpacing' (int)
     *      'color' (string) HEX color
     *      'font' (string) Arial, Times New Roman ...
     *      'fontSize' (int) 8, 9, 10, 11 ...
     *      'highlight' (string) HEX color
     *      'italic' (bool)
     *      'lang' (string)
     *      'strikethrough' (bool)
     *      'underline' (string) single
     * Options shape:
     *      'coordinateX' (int) EMUs (English Metric Unit)
     *      'coordinateY' (int) EMUs (English Metric Unit)
     *      'fillColor' (string) FF0000, 00FFFF,...
     *      'name' (string) internal name
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     * Options table:
     *      'backgroundColor' (string) HEX color
     *      'columnWidths' (int|array) column width fix (int) or column width variable (array). EMUs (English Metric Unit)
     *      'coordinateX' (int) EMUs (English Metric Unit)
     *      'coordinateY' (int) EMUs (English Metric Unit)
     *      'name' (string) internal name
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     * Options table-row:
     *      'backgroundColor' (string) HEX color
     *      'height' (int)
     * Options table-cell:
     *      'backgroundColor' (string) HEX color
     *      'textDirection' (string) horz, vert, vert270, wordArtVert, eaVert, mongolianVert, wordArtVertRtl
     *      'verticalAlign' (string) top, middle, bottom, topCentered, middleCentered, bottomCentered
     * Options table-cell-paragraph:
     *      'align' (string) left, center, right, justify, distributed
     *      'bold' (bool)
     *      'characterSpacing' (int)
     *      'color' (string) HEX color
     *      'font' (string) Arial, Times New Roman ...
     *      'fontSize' (int) 8, 9, 10, 11 ...
     *      'highlight' (string) HEX color
     *      'indentation' (int) EMUs (English Metric Unit)
     *      'italic' (bool)
     *      'lang' (string)
     *      'lineSpacing' (int|float) 1, 1.5, 2...
     *      'marginLeft' (int) EMUs (English Metric Unit) (0 >= and <= 51206400)
     *      'marginRight' (int) EMUs (English Metric Unit) (0 >= and <= 51206400)
     *      'rtl' (bool) RTL
     *      'spacingAfter' (int) points (0 >= and <= 158400)
     *      'spacingBefore' (int) points (0 >= and <= 158400)
     *      'strikethrough' (bool)
     *      'underline' (string) single
     * Options video:
     *      'coordinateX' (int) EMUs (English Metric Unit)
     *      'coordinateY' (int) EMUs (English Metric Unit)
     *      'name' (string) internal name
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     * Other options:
     *     'customAttributes' (array)
     * @throws Exception method not available
     */
    public function customizeElementStyles($referenceNode, $options = array())
    {
        if (!file_exists(__DIR__ . '/PptxCustomizer.php')) {
            PhppptxLogger::logger('This method is not available for your license.', 'fatal');
        }

        // default options
        if (!isset($referenceNode['target'])) {
            $referenceNode['target'] = 'slide';
        }
        if (!isset($referenceNode['type'])) {
            $referenceNode['type'] = 'shape';
        }

        $contentsType = $this->getContentsDOM($referenceNode['target'], $referenceNode);
        $options['tagType'] = $referenceNode['type'];

        // get the referenceNode
        $query = $this->getPptxQuery($referenceNode);
        $contentNodes = $contentsType['xpath']->query($query);

        $customizer = new PptxCustomizer();
        foreach ($contentNodes as $contentNode) {
            $customizer->customize($contentNode, $options);
        }

        // refresh contents
        $this->zipPptx->addContent($contentsType['path'], $contentsType['dom']->saveXML());

        // free DOMDocument resources
        $contentsType['dom'] = null;
    }

    /**
     * Returns the info of a PptxPath query such as number of elements and the xpath query
     *
     * @access public
     * @param array $referenceNode reference node
     *      'target' (string) slide (default)
     *      'type' (string) audio, chart, diagram, image, paragraph, run, section, shape (text box) (default), slide, table, table-row, table-cell, table-cell-paragraph, video
     *      'contains' (string) for paragraph, run, shape, table, table-row, table-cell, table-cell-paragraph types
     *      'occurrence' (int) exact occurrence (1 is the first occurrence), (array) occurrences, (string) range of contents (e.g.: 2..9, 2.., ..9) or first() or last()
     *      'attributes' (array) node or descendant attributes
     *      'parent' (string) immediate children (default as '/', any parent) or any other parent (a:tbl/, p:sp/...)
     *      'rootParent' (string) root parent. Default as p:spTree for slide elements and p:presentation for presentation elements
     *      'customQuery' (string) if set overwrites all previous references. It must be a valid XPath query
     * @return array
     * @throws Exception method not available
     */
    public function getPptxPathQueryInfo($referenceNode)
    {
        if (!file_exists(__DIR__ . '/PptxPath.php')) {
            PhppptxLogger::logger('This method is not available for your license.', 'fatal');
        }

        // default options
        if (!isset($referenceNode['target'])) {
            $referenceNode['target'] = 'slide';
        }
        if (!isset($referenceNode['type'])) {
            $referenceNode['type'] = 'shape';
        }

        $contentsType = $this->getContentsDOM($referenceNode['target'], $referenceNode);

        // get the referenceNode
        $query = $this->getPptxQuery($referenceNode);
        $contentNodes = $contentsType['xpath']->query($query);

        // free DOMDocument resources
        $contentsType['dom'] = null;

        return array(
            'elements' => $contentNodes,
            'length' => $contentNodes->length,
            'query' => $query,
        );
    }

    /**
     * Moves elements to other slides using PptxPath queries
     *
     * @access public
     * @param array $referenceNode reference node to move
     *      'target' (string) slide (default)
     *      'type' (string) audio, image, shape (text box) (default), slide, table, video
     *      'contains' (string) for shape, table types
     *      'occurrence' (int) exact occurrence (1 is the first occurrence), (array) occurrences, (string) range of contents (e.g.: 2..9, 2.., ..9) or first() or last()
     *      'attributes' (array) node or descendant attributes
     *      'parent' (string) immediate children (default as '/', any parent) or any other parent (a:tbl/, p:sp/...)
     *      'rootParent' (string) root parent. Default as p:spTree for slide elements and p:presentation for presentation elements
     *      'customQuery' (string) if set overwrites all previous references. It must be a valid XPath query
     * @param array $referenceNodeTo reference node to move the elements
     *      'target' (string) slide (default)
     *      'type' (string) slide (default)
     *      'occurrence' (int) exact occurrence (1 is the first occurrence) or first() or last(). Default as last()
     *      'attributes' (array) node or descendant attributes
     *      'parent' (string) immediate children (default as '/', any parent) or any other parent (a:tbl/, p:sp/...)
     *      'rootParent' (string) root parent. Default as p:spTree for slide elements and p:presentation for presentation elements
     *      'customQuery' (string) if set overwrites all previous references. It must be a valid XPath query
     * @throws Exception method not available
     */
    public function moveElement($referenceNode, $referenceNodeTo)
    {
        if (!file_exists(__DIR__ . '/PptxPath.php')) {
            PhppptxLogger::logger('This method is not available for your license.', 'fatal');
        }

        // default options
        if (!isset($referenceNode['target'])) {
            $referenceNode['target'] = 'slide';
        }
        if (!isset($referenceNode['type'])) {
            $referenceNode['type'] = 'shape';
        }
        if (!isset($referenceNodeTo['target'])) {
            $referenceNodeTo['target'] = 'slide';
        }
        if (!isset($referenceNodeTo['type'])) {
            $referenceNodeTo['type'] = 'slide';
        }
        if (!isset($referenceNodeTo['occurrence'])) {
            $referenceNodeTo['occurrence'] = 'last()';
        }

        $contentsType = $this->getContentsDOM($referenceNode['target'], $referenceNode);
        if ($referenceNode['type'] == 'slide') {
            $contentsTypeTo = $contentsType;
        } else {
            $contentsTypeTo = $this->getContentsDOM($referenceNodeTo['target'], $referenceNodeTo);
        }

        // get the referenceNode
        $query = $this->getPptxQuery($referenceNode);
        $contentNodes = $contentsType['xpath']->query($query);

        // get the referenceNodeTo
        $query = $this->getPptxQuery($referenceNodeTo);
        $contentNodesTo = $contentsTypeTo['xpath']->query($query);

        // get the XML contend to be moved
        $referenceContentXML = '';
        $removeNodes = [];
        foreach ($contentNodes as $contentNode) {
            $referenceContentXML .= $contentNode->ownerDocument->saveXML($contentNode);

            // remove referenceNode
            $contentNode->parentNode->removeChild($contentNode);
        }

        // create a cursor tag to be replaced by the moved content
        $cursor = $contentsType['dom']->createElement('cursor', 'CursorContent');

        if ($contentNodes->length > 0 && $contentNodesTo->length > 0) {
            if ($referenceNode['type'] == 'slide') {
                // move slide
                if (isset($referenceNodeTo['occurrence']) && ($referenceNodeTo['occurrence'] == 'first()')) {
                    // as first slide
                    $contentNodesTo->item(0)->parentNode->insertBefore($cursor, $contentNodesTo->item(0)->parentNode->firstChild);
                } else {
                    // other position
                    $contentNodesTo->item(0)->parentNode->insertBefore($cursor, $contentNodesTo->item(0)->nextSibling);
                }

                // refresh contents
                $this->zipPptx->addContent($contentsType['path'], str_replace('<cursor>CursorContent</cursor>', $referenceContentXML, $contentsType['dom']->saveXML()));
                $this->pptxPresentationDOM = $this->zipPptx->getContent('ppt/presentation.xml', 'DOMDocument');
            } else {
                // move contents to other slide

                $slideContents = $this->zipPptx->getSlides();

                // set default options
                if ($referenceNodeTo['occurrence'] == 'last()' || $referenceNodeTo['occurrence'] == '-1') {
                    $referenceNodeTo['occurrence'] = count($slideContents);
                }
                if ($referenceNodeTo['occurrence'] == 'first()') {
                    $referenceNodeTo['occurrence'] = count($slideContents);
                }

                // get the slide to add the elements moved
                $slideContent = $slideContents[$referenceNodeTo['occurrence'] - 1];

                // handle external relationships
                // check if the XML to be moved includes r:embed elements
                $newContentDOM = $this->xmlUtilities->generateDomDocument('<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">' . $referenceContentXML . '</p:sld>');
                $newContentXPath = new DOMXPath($newContentDOM);
                $newContentXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
                $newContentXPath->registerNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
                $newContentXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
                $newContentXPath->registerNamespace('p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
                $newContentXPath->registerNamespace('p15', 'http://schemas.microsoft.com/office/powerpoint/2012/main');
                $newContentXPath->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                $newContentXPath->registerNamespace('dgm', 'http://schemas.openxmlformats.org/drawingml/2006/diagram');
                $nodesEmbed = $newContentXPath->query('//*[@r:embed or @r:link or @r:id]');
                if ($nodesEmbed->length > 0) {
                    // get rels
                    $slideRelsContentsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $contentsType['path']) . '.rels';
                    $slideRelsContentsToPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $slideContent['path']) . '.rels';
                    $slideRelsContentsToContent = $this->zipPptx->getContent($slideRelsContentsToPath);
                    $slideRelsDOM = $this->zipPptx->getContent($slideRelsContentsPath, 'DOMDocument');
                    $slideRelsXPath = new DOMXPath($slideRelsDOM);
                    $slideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                    foreach ($nodesEmbed as $nodeEmbed) {
                        if ($nodeEmbed->hasAttribute('r:embed')) {
                            $originalRelsId = $nodeEmbed->getAttribute('r:embed');
                        } else if ($nodeEmbed->hasAttribute('r:link')) {
                            $originalRelsId = $nodeEmbed->getAttribute('r:link');
                        } else if ($nodeEmbed->hasAttribute('r:id')) {
                            $originalRelsId = $nodeEmbed->getAttribute('r:id');
                        }
                        if (isset($originalRelsId)) {
                            $nodesRels = $slideRelsXPath->query('//xmlns:Relationship[@Id="'.$originalRelsId.'"]');
                            if ($nodesRels->length > 0) {
                                // set a new relationship id in the content to be added
                                $relationshipId = $this->generateUniqueId();
                                if ($nodeEmbed->hasAttribute('r:embed')) {
                                    $nodeEmbed->setAttribute('r:embed', 'rId'.$relationshipId);
                                } else if ($nodeEmbed->hasAttribute('r:link')) {
                                    $nodeEmbed->setAttribute('r:link', 'rId'.$relationshipId);
                                } else if ($nodeEmbed->hasAttribute('r:id')) {
                                    $nodeEmbed->setAttribute('r:id', 'rId'.$relationshipId);
                                }

                                // add relationship to the new slide
                                $newRelationshipContent = $nodesRels->item(0)->ownerDocument->saveXML($nodesRels->item(0));
                                $newRelationshipContent = str_replace('Id="'.$originalRelsId.'"', 'Id="rId'.$relationshipId.'"', $newRelationshipContent);
                                $slideRelsContentsToContent = str_replace('</Relationships>', $newRelationshipContent . '</Relationships>', $slideRelsContentsToContent);
                            }
                        }
                    }
                    $this->zipPptx->addContent(str_replace('ppt/slides/', 'ppt/slides/_rels/', $slideContent['path']) . '.rels', $slideRelsContentsToContent);

                    // free DOMDocument resources
                    $slideRelsDOM = null;
                }

                // get slide cNvPr ids
                $idsCnvPr = array();
                $nodesCnvPr = $newContentXPath->query('//p:cNvPr');
                foreach ($nodesCnvPr as $nodeCnvPr) {
                    if ($nodeCnvPr->hasAttribute('id')) {
                        $idsCnvPr[] = $nodeCnvPr->getAttribute('id');
                    }
                }
                // generate new cNvPr ids
                $nodesId = $newContentXPath->query('//p:cNvPr[@id]');
                foreach ($nodesId as $nodeId) {
                    // clean timing tags
                    $this->cleanTimingTags($nodeId->getAttribute('id'), $contentsType['xpath'], $referenceNode['type']);

                    // generate and set a new cNvPr id and name
                    $newPlaceholderId = null;
                    while (!isset($newPlaceholderId)) {
                        $randomId = mt_rand(999, 9999999);
                        if (!in_array($randomId, $idsCnvPr)) {
                            $idsCnvPr[] = $randomId;
                            $newPlaceholderId = $randomId;
                        }
                    }
                    $nodeId->setAttribute('id', $newPlaceholderId);
                }
                // get the XML to be added without the parent tag
                $newReferenceContentXML = '';
                foreach ($newContentDOM->firstChild->childNodes as $node) {
                    $newReferenceContentXML .= $node->ownerDocument->saveXML($node);
                }

                $slideContent['content'] = str_replace('</p:spTree>', $newReferenceContentXML . '</p:spTree>', $slideContent['content']);

                // refresh contents
                // referenceNode slide
                $this->zipPptx->addContent($contentsType['path'], $contentsType['dom']->saveXML());
                // referenceNodeTo slide
                $this->zipPptx->addContent($slideContent['path'], $slideContent['content']);

                // free DOMDocument resources
                $slideContentDOM = null;
            }
        }

        // free DOMDocument resources
        $contentsType['dom'] = null;
    }

    /**
     * Removes elements using PptxPath queries
     *
     * @access public
     * @param array $referenceNode reference node to remove
     *      'target' (string) slide (default)
     *      'type' (string) audio, chart, diagram, image, paragraph, run, section, shape (text box) (default), slide, table, table-row, table-cell, table-cell-paragraph, video
     *      'contains' (string) for paragraph, run, shape, table, table-row, table-cell, table-cell-paragraph types
     *      'occurrence' (int) exact occurrence (1 is the first occurrence), (array) occurrences, (string) range of contents (e.g.: 2..9, 2.., ..9) or first() or last()
     *      'attributes' (array) node or descendant attributes
     *      'parent' (string) immediate children (default as '/', any parent) or any other parent (a:tbl/, p:sp/...)
     *      'rootParent' (string) root parent. Default as p:spTree for slide elements and p:presentation for presentation elements
     *      'customQuery' (string) if set overwrites all previous references. It must be a valid XPath query
     * @throws Exception method not available
     */
    public function removeElement($referenceNode)
    {
        if (!file_exists(__DIR__ . '/PptxPath.php')) {
            PhppptxLogger::logger('This method is not available for your license.', 'fatal');
        }

        // default options
        if (!isset($referenceNode['target'])) {
            $referenceNode['target'] = 'slide';
        }
        if (!isset($referenceNode['type'])) {
            $referenceNode['type'] = 'shape';
        }

        $contentsType = $this->getContentsDOM($referenceNode['target'], $referenceNode);

        // get the referenceNode
        $query = $this->getPptxQuery($referenceNode);
        $contentNodes = $contentsType['xpath']->query($query);

        // remove elements
        foreach ($contentNodes as $contentNode) {
            // clean related timing tags
            $nodesCNvPr = $contentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cNvPr');
            if ($nodesCNvPr->length > 0) {
                $pspId = $nodesCNvPr->item(0)->getAttribute('id');
                if (!empty($pspId)) {
                    $this->cleanTimingTags($pspId, $contentsType['xpath'], $referenceNode['type']);
                }
            }

            // fix paragraph and table cell paragraphs by adding an empty paragraph if the parent node has no paragraph
            if (($referenceNode['type'] == 'paragraph' || $referenceNode['type'] == 'table-cell-paragraph') && ($contentNode->parentNode->tagName == 'p:txBody' || $contentNode->parentNode->tagName == 'a:txBody')) {
                $pChilds = $contentNode->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'p');
                // if the parent node doesn't include a paragraph node, add it
                if ($pChilds->length > 1) {
                    $pChilds->item(0)->parentNode->removeChild($contentNode);
                } else {
                    $emptyP = $contentNode->ownerDocument->createElement('a:p');
                    $contentNode->parentNode->appendChild($emptyP);
                    $contentNode->parentNode->removeChild($contentNode);
                }
            } else {
                $contentNode->parentNode->removeChild($contentNode);
            }

            if ($referenceNode['type'] == 'section') {
                // remove p14:sectionLst if doesn't have children
                $nodesSectionLst = $contentsType['xpath']->query('//p14:sectionLst');
                if ($nodesSectionLst->length > 0 && $nodesSectionLst->item(0)->childNodes->length == 0) {
                    $nodesSectionLst->item(0)->parentNode->removeChild($nodesSectionLst->item(0));
                }
            }
        }

        // refresh contents
        $this->zipPptx->addContent($contentsType['path'], $contentsType['dom']->saveXML());
        if (in_array($referenceNode['type'], array('section', 'slide'))) {
            $this->pptxPresentationDOM = $this->zipPptx->getContent('ppt/presentation.xml', 'DOMDocument');
        }

        // free DOMDocument resources
        $contentsType['dom'] = null;
    }

    /**
     * Removes notes in slide
     *
     * @access public
     * @param array $options
     * @throws Exception position not valid
     */
    public function removeNotesSlide($options = array())
    {
        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // get notes of the slide
        $slideRelsContentsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
        $slideRelsDOM = $this->zipPptx->getContent($slideRelsContentsPath, 'DOMDocument');
        $slideRelsXPath = new DOMXPath($slideRelsDOM);
        $slideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $nodesNotesSlide = $slideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"]');
        if ($nodesNotesSlide->length > 0) {
            // keep notesSlides targets
            $targetsNotesSlide = array();
            // remove notes from the slide rels
            foreach ($nodesNotesSlide as $nodeNotesSlide) {
                $targetsNotesSlide[] = str_replace('../notesSlides', 'ppt/notesSlides', $nodeNotesSlide->getAttribute('Target'));
                $nodeNotesSlide->parentNode->removeChild($nodeNotesSlide);

            }
            // remove notes files
            foreach ($targetsNotesSlide as $targetsNoteSlide) {
                // notes file
                $this->zipPptx->deleteContent($targetsNoteSlide);
                // notes rel file
                $this->zipPptx->deleteContent(str_replace('notesSlides/', 'notesSlides/_rels/', $targetsNoteSlide) . '.rels');
            }
            // remove notes Content_Types
            $pptxContentTypesXPath = new DOMXPath($this->pptxContentTypesDOM);
            $pptxContentTypesXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');
            foreach ($targetsNotesSlide as $targetsNoteSlide) {
                $nodesOverrideTarget = $pptxContentTypesXPath->query('//xmlns:Types/xmlns:Override[@PartName="/'.$targetsNoteSlide.'"]');
                if ($nodesOverrideTarget->length > 0) {
                    foreach ($nodesOverrideTarget as $nodeOverrideTarget) {
                        $nodeOverrideTarget->parentNode->removeChild($nodeOverrideTarget);
                    }

                    // refresh contents
                    $this->zipPptx->addContent('[Content_Types].xml', $this->pptxContentTypesDOM->saveXML());
                    $this->pptxContentTypesDOM = $this->zipPptx->getContent('[Content_Types].xml', 'DOMDocument');
                }
            }

            // refresh contents
            $this->zipPptx->addContent($slideRelsContentsPath, $slideRelsDOM->saveXML());
        }

        // free DOMDocument resources
        $slideDOM = null;
        $slideRelsDOM = null;

        PhppptxLogger::logger('Remove notes.', 'info');
    }

    /**
     * Removes shape (text box) in slide
     *
     * @access public
     * @param array $position
     *      'placeholder' (array)
     *          'name' (string) name
     *          'descr' (string) alt text (descr) value
     *          'position' (int) position by order. 0 is the first order position
     *          'type' (string) title (Title), body (Body), ctrTitle (Centered Title), subTitle (Subtitle)
     * @param array $options
     * @throws Exception position not valid
     */
    public function removeShapeSlide($position, $options = array())
    {
        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // remove the pSp
        $nodePSp = $this->getPspBox($position, $slideDOM, $slideXPath, $options);
        // keep pSp ID
        $nodesCNvPr = $nodePSp->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cNvPr');
        $pspId = null;
        if ($nodesCNvPr->length > 0) {
            $pspId = $nodesCNvPr->item(0)->getAttribute('id');
        }

        $nodePSp->parentNode->removeChild($nodePSp);

        // clean related timing tags
        $this->cleanTimingTags($pspId, $slideXPath);

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;

        PhppptxLogger::logger('Remove shape slide.', 'info');
    }

    /**
     * Removes timing element
     *
     * @access public
     */
    public function removeTiming()
    {
        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // remove timing tags
        $nodesTiming = $slideXPath->query('//p:timing');
        foreach ($nodesTiming as $nodeTiming) {
            $nodeTiming->parentNode->removeChild($nodeTiming);
        }

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;

        PhppptxLogger::logger('Remove timing.', 'info');
    }

    /**
     * Generates the new PPTX file
     *
     * @access public
     * @param string $fileName path to the resulting PPTX
     * @throws Exception license is not valid
     * @throws Exception PPTX can't be saved
     * @return PptxStructure
     */
    public function savePptx($fileName = 'presentation')
    {
        PhppptxLogger::logger('Set PPTX name to: ' . $fileName . '.', 'info');

        // refresh contents
        $this->zipPptx->addContent('[Content_Types].xml', $this->pptxContentTypesDOM->saveXML());
        $this->zipPptx->addContent('ppt/_rels/presentation.xml.rels', $this->pptxRelsPresentationDOM->saveXML());
        $this->zipPptx->addContent('ppt/presentation.xml', $this->pptxPresentationDOM->saveXML());

        // delete XLSX tempfiles (charts)
        foreach ($this->tempFileXLSX as $file) {
            unlink($file);
            unlink($file . '.pptx');
        }

        PhppptxLogger::logger('Create PPTX.', 'info');

        return $this->zipPptx->savePptx($fileName);
    }

    /**
     * Generate and download a new PPTX file
     *
     * @access public
     * @param string $fileName file name
     * @param bool $removeAfterDownload remove the file after download it
     * @throws Exception PPTX can't be saved
     */
    public function savePptxAndDownload($fileName, $removeAfterDownload = false)
    {
        try {
            $this->savePptx($fileName);
        } catch (Exception $e) {
            PhppptxLogger::logger('Error while trying to write to ' . $fileName . ' . Check write access.', 'fatal');
        }

        if (!empty($fileName)) {
            $fileName = str_replace(array('/', '\\'), DIRECTORY_SEPARATOR, $fileName);
            $completeName = explode(DIRECTORY_SEPARATOR, $fileName);
            $fileNameDownload = array_pop($completeName);
        } else {
            $fileName = 'presentation';
            $fileNameDownload = 'presentation';
        }

        // check if the path has as extension, and remove it if true
        if (substr($fileNameDownload, -5) == '.pptx') {
            $fileNameDownload = substr($fileNameDownload, 0, -5);
        }

        // get absolute path to the file to be used with filesize and readfile methods
        $fileInfo = pathinfo($fileName);
        $filePath = $fileInfo['dirname'] . '/' . $fileNameDownload;

        $extension = 'pptx';
        PhppptxLogger::logger('Download file ' . $fileNameDownload . '.' . $extension . '.', 'info');
        header('Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation');
        header('Content-Disposition: attachment; filename="' . $fileNameDownload . '.' . $extension . '"');
        header('Content-Transfer-Encoding: binary');
        header('Content-Length: ' . filesize($filePath . '.' . $extension));
        readfile($filePath . '.' . $extension);

        // remove the generated file
        if ($removeAfterDownload) {
            unlink($filePath . '.' . $extension);
        }
    }

    /**
     * Sets the active slide
     *
     * @access public
     * @param array $options
     *      'position' (int) slide position. 0 is the first slide. -1 can be used to choose the last slide
     * @throws Exception position isn't set
     * @throws Exception position doesn't exist
     */
    public function setActiveSlide($options)
    {
        // if no position is set, throw an Exception
        if (!isset($options['position'])) {
            PhppptxLogger::logger('Choose a position to set the active slide.', 'fatal');
        }

        // get slides
        $slidesContents = $this->zipPptx->getSlides();

        if (isset($options['position'])) {
            // handle last position
            if ($options['position'] == -1) {
                $options['position'] = count($slidesContents) - 1;
            }

            if (!isset($slidesContents[$options['position']])) {
                PhppptxLogger::logger('The slide position doesn\'t exist.', 'fatal');
            } else {
                $this->activeSlide = array(
                    'position' => (int)$options['position'],
                );
            }
        }

        PhppptxLogger::logger('Set active slide.', 'info');
    }

    /**
     * Sets the background color
     *
     * @access public
     * @param string $color hexadecimal color value without # : FFFF00, 92D050...
     * @param array $options
     */
    public function setBackgroundColor($color, $options = array())
    {
        // clean # from color value
        $color = str_replace('#', '', $color);

        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);

        // remove p:cSld/p:bg tag if any exists
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $nodesBgToRemove = $slideXPath->query('//p:cSld/p:bg');
        if ($nodesBgToRemove->length > 0) {
            foreach ($nodesBgToRemove as $nodeBgToRemove) {
                $nodeBgToRemove->parentNode->removeChild($nodeBgToRemove);
            }
        }

        // add the new p:bg tag before p:spTree
        $newbgXML = '<p:bg xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:bgPr><a:solidFill><a:srgbClr val="'.$color.'"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
        $newBgFragment = $slideDOM->createDocumentFragment();
        $newBgFragment->appendXML($newbgXML);

        $nodesSpTree = $slideXPath->query('//p:spTree');
        if ($nodesSpTree->length > 0) {
            $nodesSpTree->item(0)->parentNode->insertBefore($newBgFragment, $nodesSpTree->item(0));
        }

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;

        PhppptxLogger::logger('Set background color.', 'info');
    }

    /**
     * Marks the document as final
     *
     * @access public
     */
    public function setMarkAsFinal()
    {
        $this->addProperties(array('contentStatus' => 'Final'));
        $this->generateOverride('/docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml');

        PhppptxLogger::logger('Enable mark as final.', 'info');
    }

    /**
     * Sets presentation settings
     *
     * @access public
     * @param array $options
     *      'height' (int): measurement in EMUs (914400 >= and <= 51206400)
     *      'notes' (array): notes slide properties
     *              'height' (int)
     *              'width' (int)
     *      'readOnly' (bool) set as read only
     *      'rtl' (bool) set to true for right to left
     *      'type' (string): 35mm, 35mm-portrait, A3, A3-portrait, A4, A4-portrait, B4ISO, B4ISO-portrait, B5ISO, B5ISO-portrait, banner, banner-portrait, ledger, ledger-portrait, letter, letter-portrait, overhead, overhead-portrait, screen16x10, screen16x10-portrait, screen16x9, screen16x9-portrait, screen4x3, screen4x3-portrait, standard, standard-portrait, widescreen, widescreen-portrait, custom
     *      'width' (int): measurement in EMUs (914400 >= and <= 51206400)
     */
    public function setPresentationSettings($options = array())
    {
        $referenceTypes = array(
            '35mm' => array(
                'height' => 6858000,
                'width' => 10287000,
                'type' => '35mm',
            ),
            '35mm-portrait' => array(
                'height' => 10287000,
                'width' => 6858000,
                'type' => '35mm',
            ),
            'A3' => array(
                'height' => 9601200,
                'width' => 12801600,
                'type' => 'A3',
            ),
            'A3-portrait' => array(
                'height' => 12801600,
                'width' => 9601200,
                'type' => 'A3',
            ),
            'A4' => array(
                'height' => 6858000,
                'width' => 9906000,
                'type' => 'A4',
            ),
            'A4-portrait' => array(
                'height' => 9906000,
                'width' => 6858000,
                'type' => 'A4',
            ),
            'B4ISO' => array(
                'height' => 8120063,
                'width' => 10826750,
                'type' => 'B4ISO',
            ),
            'B4ISO-portrait' => array(
                'height' => 10826750,
                'width' => 8120063,
                'type' => 'B4ISO',
            ),
            'B5ISO' => array(
                'height' => 5376863,
                'width' => 7169150,
                'type' => 'B5ISO',
            ),
            'B5ISO-portrait' => array(
                'height' => 7169150,
                'width' => 5376863,
                'type' => 'B5ISO',
            ),
            'banner' => array(
                'height' => 914400,
                'width' => 7315200,
                'type' => 'banner',
            ),
            'banner-portrait' => array(
                'height' => 7315200,
                'width' => 914400,
                'type' => 'banner',
            ),
            'custom' => array(
                'height' => 6858000,
                'width' => 12192000,
                'type' => 'custom',
            ),
            'ledger' => array(
                'height' => 9134475,
                'width' => 12179300,
                'type' => 'ledger',
            ),
            'ledger-portrait' => array(
                'height' => 12179300,
                'width' => 9134475,
                'type' => 'ledger',
            ),
            'letter' => array(
                'height' => 6858000,
                'width' => 9144000,
                'type' => 'letter',
            ),
            'letter-portrait' => array(
                'height' => 9144000,
                'width' => 6858000,
                'type' => 'letter',
            ),
            'overhead' => array(
                'height' => 6858000,
                'width' => 9144000,
                'type' => 'overhead',
            ),
            'overhead-portrait' => array(
                'height' => 9144000,
                'width' => 6858000,
                'type' => 'overhead',
            ),
            'screen16x10' => array(
                'height' => 5715000,
                'width' => 9144000,
                'type' => 'screen16x10',
            ),
            'screen16x10-portrait' => array(
                'height' => 9144000,
                'width' => 5715000,
                'type' => 'screen16x10',
            ),
            'screen16x9' => array(
                'height' => 5143500,
                'width' => 9144000,
                'type' => 'screen16x9',
            ),
            'screen16x9-portrait' => array(
                'height' => 9144000,
                'width' => 5143500,
                'type' => 'screen16x9',
            ),
            'screen4x3' => array(
                'height' => 6858000,
                'width' => 9144000,
                'type' => 'screen4x3',
            ),
            'screen4x3-portrait' => array(
                'height' => 9144000,
                'width' => 6858000,
                'type' => 'screen4x3',
            ),
            'widescreen16x9' => array(
                'height' => 6858000,
                'width' => 12192000,
            ),
            'widescreen16x9-portrait' => array(
                'height' => 12192000,
                'width' => 6858000,
            ),
        );

        // handle height option
        if ((isset($options['type']) && array_key_exists($options['type'], $referenceTypes)) || isset($options['height'])) {
            if (isset($options['type']) && array_key_exists($options['type'], $referenceTypes)) {
                $heightValue = $referenceTypes[$options['type']]['height'];
            }
            if (isset($options['height'])) {
                $heightValue = $options['height'];
            }
            $nodesSldSz = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldSz');
            if ($nodesSldSz->length > 0) {
                if (isset($heightValue)) {
                    if ($heightValue < 914400) {
                        $heightValue = 914400;
                        PhppptxLogger::logger('The allowed minimum height value is 914400.', 'warning');
                    }
                    if ($heightValue > 51206400) {
                        $heightValue = 51206400;
                        PhppptxLogger::logger('The allowed maximum height value is 51206400.', 'warning');
                    }

                    $nodesSldSz->item(0)->setAttribute('cy', $heightValue);

                    if (isset($referenceTypes[$options['type']]['type'])) {
                        $nodesSldSz->item(0)->setAttribute('type', $referenceTypes[$options['type']]['type']);
                    }
                }
            }
        }

        // handle notes options
        if (isset($options['notes'])) {
            // height
            if (isset($options['notes']['height'])) {
                $nodesNotesSz = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'notesSz');
                if ($nodesNotesSz->length > 0) {
                    $nodesNotesSz->item(0)->setAttribute('cy', $options['notes']['height']);
                }
            }

            // width
            if (isset($options['notes']['width'])) {
                $nodesNotesSz = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'notesSz');
                if ($nodesNotesSz->length > 0) {
                    $nodesNotesSz->item(0)->setAttribute('cx', $options['notes']['width']);
                }
            }
        }

        // handle rtl option
        if (self::$rtl && !isset($options['rtl'])) {
            $options['rtl'] = true;
        }
        if (isset($options['rtl'])) {
            $nodesPresentation = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'presentation');
            if ($nodesPresentation->length > 0) {
                if ($options['rtl']) {
                    $nodesPresentation->item(0)->setAttribute('rtl', '1');
                } else if ($nodesPresentation->item(0)->hasAttribute('rtl')) {
                    $nodesPresentation->item(0)->removeAttribute('rtl');
                }
            }
        }

        // handle width option
        if ((isset($options['type']) && array_key_exists($options['type'], $referenceTypes)) || isset($options['width'])) {
            if (isset($options['type']) && array_key_exists($options['type'], $referenceTypes)) {
                $widthValue = $referenceTypes[$options['type']]['width'];
            }
            if (isset($options['width'])) {
                $widthValue = $options['width'];
            }
            $nodesSldSz = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldSz');
            if ($nodesSldSz->length > 0) {
                if (isset($widthValue)) {
                    if ($widthValue < 914400) {
                        $widthValue = 914400;
                        PhppptxLogger::logger('The allowed minimum width value is 914400.', 'warning');
                    }
                    if ($widthValue > 51206400) {
                        $widthValue = 51206400;
                        PhppptxLogger::logger('The allowed maximum width value is 51206400.', 'warning');
                    }

                    $nodesSldSz->item(0)->setAttribute('cx', $widthValue);

                    if (isset($referenceTypes[$options['type']]['type'])) {
                        $nodesSldSz->item(0)->setAttribute('type', $referenceTypes[$options['type']]['type']);
                    }
                }
            }
        }

        // handle readOnly option
        if (isset($options['readOnly'])) {
            // property set in presProps.xml
            $presPropsDOM = $this->zipPptx->getContent('ppt/presProps.xml', 'DOMDocument');
            $nodesReadonlyRecommended = $presPropsDOM->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2017/10/main', 'readonlyRecommended');

            if ($options['readOnly']) {
                if ($nodesReadonlyRecommended->length > 0) {
                    // enable readOnly in the existing option
                    $nodesReadonlyRecommended->item(0)->setAttribute('val', '1');
                } else {
                    // the option doesn't exist. Create and add it
                    $newNodeExtContent = '<p:ext uri="{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p1710:readonlyRecommended val="1" xmlns:p1710="http://schemas.microsoft.com/office/powerpoint/2017/10/main"/></p:ext>';
                    $newNodeExt = $presPropsDOM->createDocumentFragment();
                    $newNodeExt->appendXML($newNodeExtContent);

                    $presPropsDOM->documentElement->firstChild->appendChild($newNodeExt);
                }
            } else {
                if ($nodesReadonlyRecommended->length > 0) {
                    // disable readOnly in the existing option.
                    // If the option doesn't exist, it's not applied as default
                    $nodesReadonlyRecommended->item(0)->setAttribute('val', '0');
                }
            }

            // refresh contents
            $this->zipPptx->addContent('ppt/presProps.xml', $presPropsDOM->saveXML());
        }

        // refresh contents
        $this->zipPptx->addContent('ppt/presentation.xml', $this->pptxPresentationDOM->saveXML());
        $this->pptxPresentationDOM = $this->zipPptx->getContent('ppt/presentation.xml', 'DOMDocument');

        PhppptxLogger::logger('Set presentation settings.', 'info');
    }

    /**
     * Sets global right to left options
     * @access public
     * @param array $options
     *      'rtl' (bool)
     * @return void
     */
    public function setRtl($options = array('rtl' => true))
    {
        if (isset($options['rtl']) && $options['rtl']) {
            self::$rtl = true;
        }

        PhppptxLogger::logger('Enable RTL mode.', 'info');
    }

    /**
     * Sets slide settings
     *
     * @access public
     * @param array $options
     *      'layout' (string) layout (name) to be applied to the slide
     *      'show' (bool) show or hide the slide
     * @throws Exception layout name doesn't exist
     */
    public function setSlideSettings($options = array())
    {
        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);

        // layout
        if (isset($options['layout'])) {
            // get the slide rels
            $slideRelsContentsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
            $slideRelsDOM = $this->zipPptx->getContent($slideRelsContentsPath, 'DOMDocument');

            // check if the layout name exists
            $foundLayout = false;
            $layoutsContents = $this->zipPptx->getLayouts();
            $newLayoutPath = '';
            foreach ($layoutsContents as $layoutContents) {
                if ($layoutContents['name'] == $options['layout']) {
                    $foundLayout = true;

                    $newLayoutPath = $layoutContents['path'];
                }
            }

            if (!$foundLayout) {
                PhppptxLogger::logger('The chosen layout name \'' . $options['layout'] . '\' doesn\'t exist. Choose a valid layout.', 'fatal');
            }

            // correct layout relative path
            $newLayoutPath = str_replace('ppt/', '../', $newLayoutPath);

            // add the new layout in the rels file
            $slideRelsXPath = new DOMXPath($slideRelsDOM);
            $slideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
            $nodesSlideLayout = $slideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"]');
            if ($nodesSlideLayout->length > 0) {
                $nodesSlideLayout->item(0)->setAttribute('Target', $newLayoutPath);
            }

            // refresh contents
            $this->zipPptx->addContent($slideRelsContentsPath, $slideRelsDOM->saveXML());

            // free DOMDocument resources
            $slideRelsDOM = null;
        }

        // show
        if (isset($options['show'])) {
            $nodesSld = $slideDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sld');
            if ($nodesSld->length > 0) {
                if ($options['show']) {
                    $nodesSld->item(0)->setAttribute('show', '1');
                } else {
                    $nodesSld->item(0)->setAttribute('show', '0');
                }
            }
        }

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;

        PhppptxLogger::logger('Set slide settings.', 'info');
    }

    /**
     * Transforms files
     *
     * libreoffice method supports:
     *     PPTX to PDF, PPT, ODP
     *     PPT to PPTX, PDF, ODP
     *     ODP to PPTX, PDF, PPT
     *
     * mspowerpoint method supports:
     *     PPTX to PDF, PPT
     *     PPT to PDF, PPTX
     *
     * @access public
     * @param string $source
     * @param string $target
     * @param string $method libreoffice, mspowerpoint
     * @param array $options
     * libreoffice method options:
     *   'debug' (bool) false (default) or true. Shows debug information about the plugin conversion
     *   'escapeshellarg' (bool) false (default) or true. Applies escapeshellarg to escape source and LibreOffice path strings
     *   'extraOptions' (string) extra parameters to be used when doing the conversion
     *   'homeFolder' (string) set a custom home folder to be used for the conversions
     *   'outdir' (string) set the outdir path. Useful when the PDF output path is not the same than the running script
     *   'path' (string) set the path to LibreOffice. If set, overwrite the path option in phppptxconfig.ini
     *
     * mspowerpoint method options
     * @throws Exception method not available in license
     * @throws Exception unsupported file type
     * @throws Exception PHP COM extension is not available
     */
    public function transform($source, $target, $method = null, $options = array())
    {
        if (file_exists(__DIR__ . '/TransformPlugin.php')) {
            if (isset($this->phppptxconfig['transform']['method']) && $method === null) {
                $method = $this->phppptxconfig['transform']['method'];
            }

            switch ($method) {
                case 'mspowerpoint':
                    $convert = new TransformMSPowerPoint();
                    $convert->transform($source, $target, $options);
                    break;
                case 'libreoffice':
                default:
                    $convert = new TransformLibreOffice();
                    $convert->transform($source, $target, $options);
                    break;
            }
        } else {
            PhppptxLogger::logger('This method is not available for your license.', 'fatal');
        }
    }

    /**
     * Generates a unique number not used in elements
     *
     * @access public
     * @static
     * @param int $min
     * @param int $max
     * @return int
     */
    public static function uniqueNumberId($min, $max)
    {
        $proposedId = mt_rand($min, $max);
        if (in_array($proposedId, self::$elementsId)) {
            $proposedId = self::uniqueNumberId($min, $max);
        }
        self::$elementsId[] = $proposedId;

        PhppptxLogger::logger('New ID: ' . $proposedId, 'debug');

        return $proposedId;
    }

    /**
     * Adds a media
     *
     * @param array $media
     * @param array $position
     *      'new' (array) a new position is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     * @param array $mediaStyles
     *      'image' (array)
     *          'image' (string) image to be used as preview. Set a default one if not set
     *          'mime' (string) forces a mime (image/jpg, image/jpeg, image/png, image/gif, image/webp)
     *      'mime' (string) forces a mime (audio/mpeg, audio/x-wav, audio/x-ms-wma, audio/unknown)
     * @param array $options
     *      'descr' (string) alt text (descr) value
     *      'type' (string) audio, video
     * @throws Exception media doesn't exist
     * @throws Exception media format is not supported
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     * @throws Exception position not valid
     */
    protected function addMedia($media, $position, $mediaStyles = array(), $options = array())
    {
        // get the internal active slide
        $slideContents = $this->zipPptx->getSlides();
        $activeSlideContent = $slideContents[$this->activeSlide['position']];
        $slideDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
        $slideXPath = new DOMXPath($slideDOM);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        $idMedia2006 = $this->generateUniqueId();
        $mediaStyles['rIdMedia2006'] = $idMedia2006;
        $idMedia2007 = $this->generateUniqueId();
        $mediaStyles['rIdMedia2007'] = $idMedia2007;
        $idImage = $this->generateUniqueId();
        $mediaStyles['rIdImage'] = $idImage;

        // create and add the new media
        $mediaElement = new CreatePic();
        $mediaElement->addElementMedia($slideDOM, $position, $mediaStyles, $options);

        // generate and add the new relationships and content types

        // media2006 and media2007. The same target for both
        $newRelationshipMedia2006 = '<Relationship Id="rId' . $idMedia2006 . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'.$options['type'].'" Target="../media/'.$options['type'].$idMedia2006.'.'.$mediaStyles['contents']['extension'].'" />';
        $newRelationshipMedia2007 = '<Relationship Id="rId' . $idMedia2007 . '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/'.$options['type'].''.$idMedia2006.'.'.$mediaStyles['contents']['extension'].'" />';

        // generate content type if it does not exist yet
        $this->generateDefault($mediaStyles['contents']['extension'], $mediaStyles['contents']['mime']);

        // image
        $imageContent = '';
        if (isset($mediaStyles['image']) && isset($mediaStyles['image']['image'])) {
            $imageInformation = new ImageUtilities();
            $imageContents = $imageInformation->returnImageContents($mediaStyles['image']['image'], $options);

            $imageContent = $imageContents['content'];
            $imageExtension = $imageContents['extension'];
            $imageMimeType = $imageContents['mime'];
        } else {
            // default image
            if ($options['type'] == 'audio') {
                $imageContent = base64_decode('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');
            } else if ($options['type'] == 'video') {
                $imageContent = base64_decode('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');
            }
            $imageExtension = 'png';
            $imageMimeType = 'image/png';
        }
        $newRelationshipImage = '<Relationship Id="rId'.$idImage.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img'.$idImage.'.'.$imageExtension.'" />';
        // generate content type if it does not exist yet
        $this->generateDefault($imageExtension, $imageMimeType);

        // add the new relationships
        $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
        $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
        $this->generateRelationship($slideRelsDOM, $newRelationshipMedia2006);
        $this->generateRelationship($slideRelsDOM, $newRelationshipMedia2007);
        $this->generateRelationship($slideRelsDOM, $newRelationshipImage);

        // copy the contents with the new names. Only one file copy for both 2006 and 2007
        $this->zipPptx->addContent('ppt/media/'.$options['type'].$idMedia2006.'.'.$mediaStyles['contents']['extension'], $mediaStyles['contents']['content']);
        $this->zipPptx->addContent('ppt/media/img'.$idImage.'.'.$imageExtension, $imageContent);

        // refresh contents
        $this->zipPptx->addContent($activeSlideContent['path'], $slideDOM->saveXML());
        $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

        // free DOMDocument resources
        $slideDOM = null;
        $slideRelsDOM = null;
    }

    /**
     * Cleans p:sp box to insert new contents
     *
     * @access protected
     * @param DOMNode $nodeTxBody
     * @param DOMXPath $slideXPath
     * @param array $options
     *      'insertMode' (string) append, replace. If a content exists in the text box, handle how new contents are added. Default as append
     */
    protected function cleanPspBox($nodeTxBody, $slideXPath, $options)
    {
        if ($options['insertMode'] == 'append') {
            // remove the default empty paragraph included in the text box
            $nodesP = $nodeTxBody->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'p');
            if ($nodesP->length == 1 && empty($nodesP->item(0)->textContent)) {
                $nodesP->item(0)->parentNode->removeChild($nodesP->item(0));
            }
        } else if ($options['insertMode'] == 'replace') {
            // remove existing paragraphs
            $nodesP = $slideXPath->query('.//a:p', $nodeTxBody);
            foreach ($nodesP as $nodeP) {
                $nodeP->parentNode->removeChild($nodeP);
            }
        }
    }

    /**
     * Cleans timing tags from a specific shape ID
     *
     * @param string $id
     * @param DOMXPath $xPath
     * @param string $type audio, image, video
     */
    protected function cleanTimingTags($id, $xPath, $type = null) {
        if (!is_null($id)) {
            if ($type == 'audio') {
                // audio nodes
                $nodesToBeRemoved = $xPath->query('//p:timing//p:audio[.//p:spTgt[@spid="'.$id.'"]]');
                foreach ($nodesToBeRemoved as $nodeToBeRemoved) {
                    $nodeToBeRemoved->parentNode->removeChild($nodeToBeRemoved);
                }
            }
            if ($type == 'video') {
                // video nodes
                $nodesToBeRemoved = $xPath->query('//p:timing//p:video[.//p:spTgt[@spid="'.$id.'"]]');
                foreach ($nodesToBeRemoved as $nodeToBeRemoved) {
                    $nodeToBeRemoved->parentNode->removeChild($nodeToBeRemoved);
                }
            }

            // p:childTnLst nodes
            $nodesToBeRemoved = $xPath->query('//p:timing//p:spTgt[@spid="'.$id.'"]/ancestor::p:par[3]');
            foreach ($nodesToBeRemoved as $nodeToBeRemoved) {
                $nodeToBeRemoved->parentNode->removeChild($nodeToBeRemoved);
            }

            // p:seq nodes
            $nodesToBeRemoved = $xPath->query('//p:timing//p:spTgt[@spid="'.$id.'"]/ancestor::p:seq[1]');
            foreach ($nodesToBeRemoved as $nodeToBeRemoved) {
                $nodeToBeRemoved->parentNode->removeChild($nodeToBeRemoved);
            }

            // p:bldP nodes in p:bldLst
            $nodesToBeRemoved = $xPath->query('//p:timing/p:bldLst/p:bldP[@spid="'.$id.'"]');
            foreach ($nodesToBeRemoved as $nodeToBeRemoved) {
                $nodeToBeRemoved->parentNode->removeChild($nodeToBeRemoved);
            }

            // if p:bldLst doesn't have children, remove it
            $nodesPBldLst = $xPath->query('//p:timing/p:bldLst');
            if ($nodesPBldLst->length > 0 && $nodesPBldLst->item(0)->childNodes->length == 0) {
                $nodesPBldLst->item(0)->parentNode->removeChild($nodesPBldLst->item(0));
            }

            // remove remaing p:timing nodes without p:spTgt nodes
            $nodesToBeRemoved = $xPath->query('//p:timing[not(.//p:spTgt)]');
            foreach ($nodesToBeRemoved as $nodeToBeRemoved) {
                $nodeToBeRemoved->parentNode->removeChild($nodeToBeRemoved);
            }
        }
    }

    /**
     * Generates DEFAULT
     *
     * @access protected
     */
    protected function generateDefault($extension, $contentType)
    {
        $strContent = $this->pptxContentTypesDOM->saveXML();
        if (
            stripos($strContent, 'Extension="' . strtolower($extension) . '"') === false
        ) {
            $strContentTypes = '<Default Extension="' . $extension . '" ContentType="' . $contentType . '"></Default>';
            $defaultFragment = $this->pptxContentTypesDOM->createDocumentFragment();
            $defaultFragment->appendXML($strContentTypes);
            $this->pptxContentTypesDOM->documentElement->appendChild($defaultFragment);

            // refresh contents
            $this->zipPptx->addContent('[Content_Types].xml', $this->pptxContentTypesDOM->saveXML());
            $this->pptxContentTypesDOM = $this->zipPptx->getContent('[Content_Types].xml', 'DOMDocument');
        }
    }

    /**
     * Generates override
     *
     * @access protected
     * @param string $partName
     * @param string $contentType
     */
    protected function generateOverride($partName, $contentType)
    {
        $strContent = $this->pptxContentTypesDOM->saveXML();
        if (
                strpos($strContent, 'PartName="' . $partName . '"') === false
        ) {
            $strContentTypes = '<Override PartName="' . $partName . '" ContentType="' . $contentType . '" />';
            $overrideFragment = $this->pptxContentTypesDOM->createDocumentFragment();
            $overrideFragment->appendXML($strContentTypes);
            $this->pptxContentTypesDOM->documentElement->appendChild($overrideFragment);

            // refresh contents
            $this->zipPptx->addContent('[Content_Types].xml', $this->pptxContentTypesDOM->saveXML());
            $this->pptxContentTypesDOM = $this->zipPptx->getContent('[Content_Types].xml', 'DOMDocument');
        }
    }

    /**
     * Generate relationship
     *
     * @access protected
     * @param DOMDocument $relsDOM
     * @param string $newRelationship
     */
    protected function generateRelationship($relsDOM, $newRelationship)
    {
        $newNodeRelationship = $relsDOM->createDocumentFragment();
        $newNodeRelationship->appendXML($newRelationship);
        $relsDOM->documentElement->appendChild($newNodeRelationship);
    }

    /**
     * Generates uniqueID
     *
     * @access protected
     * @return string
     */
    protected function generateUniqueId() {
        $uniqueId = uniqid((string)mt_rand(999, 9999));

        return $uniqueId;
    }

    /**
     * Returns a DOM content based on the type
     *
     * @access protected
     * @param string $target Target presentation, slide (default), slideLayout
     * @param array $referenceNode Reference node
     * @return array DOM
     */
    protected function getContentsDOM($target = 'slide', $referenceNode = array())
    {
        if ($referenceNode['type'] == 'section' || $referenceNode['type'] == 'slide') {
            $target = 'presentation';
        }
        $contentDOM = null;
        $path = null;

        if ($target == 'presentation') {
            // get the presentation content
            $contentDOM = $this->pptxPresentationDOM;
            $path = 'ppt/presentation.xml';
        } else if ($target == 'slideLayout') {
            // get the slide layouts content of the internal active slide
            $slideContents = $this->zipPptx->getSlides();
            $activeSlideContent = $slideContents[$this->activeSlide['position']];

            // get related layout
            $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
            $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
            $contentSlideRelsXPath = new DOMXPath($slideRelsDOM);
            $contentSlideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
            $nodesSlideLayout = $contentSlideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"]');
            if ($nodesSlideLayout->length > 0) {
                if ($nodesSlideLayout->item(0)->hasAttribute('Target')) {
                    $targetSlideLayout = $nodesSlideLayout->item(0)->getAttribute('Target');
                }
                $nameSlideLayout = '';
                if (!empty($targetSlideLayout)) {
                    $path = str_replace('../', 'ppt/', $targetSlideLayout);
                    $contentSlideLayout = $this->zipPptx->getContent($path);

                    $contentDOM = $this->xmlUtilities->generateDomDocument($contentSlideLayout);
                }
            }

            // free DOMDocument resources
            $slideRelsDOM = null;
        } else {
            // get the internal active slide
            $slideContents = $this->zipPptx->getSlides();
            $activeSlideContent = $slideContents[$this->activeSlide['position']];
            $contentDOM = $this->xmlUtilities->generateDomDocument($activeSlideContent['content']);
            $path = $activeSlideContent['path'];
        }

        $contentXPath = new DOMXPath($contentDOM);
        $contentXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $contentXPath->registerNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
        $contentXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $contentXPath->registerNamespace('p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
        $contentXPath->registerNamespace('p15', 'http://schemas.microsoft.com/office/powerpoint/2012/main');
        $contentXPath->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $contentXPath->registerNamespace('dgm', 'http://schemas.openxmlformats.org/drawingml/2006/diagram');

        return array(
            'dom' => $contentDOM,
            'path' => $path,
            'xpath' => $contentXPath,
        );
    }

    /**
     * Gets extension from mime
     *
     * @access protected
     * @param string $mime
     * @return string
     */
    protected function getExtensionFromMime($mime)
    {
        $extension = '';

        switch ($mime) {
            case 'audio/mpeg':
                $extension = 'mp3';
                break;
            case 'audio/unknown':
                $extension = 'flac';
                break;
            case 'audio/x-ms-wma':
                $extension = 'wma';
                break;
            case 'audio/x-wav':
                $extension = 'wav';
                break;
            case 'image/bmp':
                $extension = 'bmp';
                break;
            case 'image/gif':
                $extension = 'gif';
                break;
            case 'image/jpg':
            case 'image/jpeg':
                $extension = 'jpeg';
                break;
            case 'image/png':
                $extension = 'png';
                break;
            case 'image/webp':
                $extension = 'webp';
                break;
            case 'video/mp4':
                $extension = 'mp4';
                break;
            case 'video/unknown':
                $extension = 'mkv';
                break;
            case 'video/x-ms-wmv':
                $extension = 'wmv';
                break;
            case 'video/x-msvideo':
                $extension = 'avi';
                break;
            default:
                break;
        }

        return strtolower($extension);
    }

    /**
     * Gets max slide id
     *
     * @return int
     */
    protected function getMaxSlideId() {
        $newIdSlide = 1;

        $sldIdLstTags = $this->pptxPresentationDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldIdLst');
        if ($sldIdLstTags->length > 0) {
            // get max slide ids
            $sldIdTags = $sldIdLstTags->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldId');
            $newIdSlide = 1;
            if ($sldIdTags->length > 0) {
                // generate a new ID from existing values
                foreach ($sldIdTags as $sldIdTag) {
                    if ($sldIdTag->hasAttribute('id')) {
                        if ((int)$sldIdTag->getAttribute('id') > $newIdSlide) {
                            $newIdSlide = (int)$sldIdTag->getAttribute('id');
                        }
                    }
                }
            }
        }

        return $newIdSlide;
    }

    /**
     * Gets mime from extension
     *
     * @access protected
     * @param string $extension
     * @return string
     */
    protected function getMimeFromExtension($extension)
    {
        $mime = '';

        switch ($extension) {
            case 'avi':
                $mime = 'video/x-msvideo';
                break;
            case 'bmp':
                $mime = 'image/bmp';
                break;
            case 'flac':
                $mime = 'audio/unknown';
                break;
            case 'gif':
                $mime = 'image/gif';
                break;
            case 'jpg':
                $mime = 'image/jpg';
                break;
            case 'jpeg':
                $mime = 'image/jpeg';
                break;
            case 'mp3':
                $mime = 'audio/mpeg';
                break;
            case 'mkv':
                $mime = 'video/unknown';
                break;
            case 'mp4':
                $mime = 'video/mp4';
                break;
            case 'png':
                $mime = 'image/png';
                break;
            case 'webp':
                $mime = 'image/webp';
                break;
            case 'wav':
                $mime = 'audio/x-wav';
                break;
            case 'wma':
                $mime = 'audio/x-ms-wma';
                break;
            case 'wmv':
                $mime = 'video/x-ms-wmv';
                break;
            default:
                break;
        }

        return strtolower($mime);
    }

    /**
     * Returns a content query based on the reference
     *
     * @access protected
     * @param array $referenceNode
     * @return string
     * @throws Exception method not available
     */
    protected function getPptxQuery($referenceNode)
    {
        if (!file_exists(__DIR__ . '/PptxPath.php')) {
            PhppptxLogger::logger('This method is not available for your license.', 'fatal');
        }

        if (isset($referenceNode['customQuery']) && !empty($referenceNode['customQuery'])) {
            $query = $referenceNode['customQuery'];
        } else {
            $query = PptxPath::xpathContentQuery($referenceNode['type'], $referenceNode);
        }
        PhppptxLogger::logger('PptxPath query: ' . $query, 'debug');

        return $query;
    }

    /**
     * Returns p:sp content based on the position or generating a new content
     *
     * @access protected
     * @param array $position
     *      'placeholder' (array) the content is added in a placeholder of the layout. One of the following options can be used to get the text box
     *          'name' (string) name
     *          'descr' (string) alt text (descr) value
     *          'position' (int) position by order. 0 is the first order position
     *          'type' (string) title (Title), body (Body), ctrTitle (Centered Title), subTitle (Subtitle)
     *      'new' (array) a new position is generated
     *          'coordinateX' (int) EMUs (English Metric Unit)
     *          'coordinateY' (int) EMUs (English Metric Unit)
     *          'sizeX' (int) EMUs (English Metric Unit)
     *          'sizeY' (int) EMUs (English Metric Unit)
     *          'name' (string) internal name. If not set, a random name is generated
     *          'order' (int) set the display order. Default after existing contents. 0 is the first order position. If the order position doesn't exist add after existing contents
     *          'textBoxStyles' (array) @see addTextBox
     * @param DOMDocument $slideDOM
     * @param DOMXPath $slideXPath
     * @param array $options
     * @return DOMNode
     * @throws Exception position not valid
     */
    protected function getPspBox($position, $slideDOM, $slideXPath, $options = array())
    {
        // get the position to add the new content
        $nodePSp = null;
        $positionSet = null; // used to debug the position
        // by placeholder attributes
        if (isset($position['placeholder']) && isset($position['placeholder']['name'])) {
            // get placeholder by name
            $nodesPSp = $slideXPath->query('//p:sp[.//p:cNvPr[@name="'.$position['placeholder']['name'].'"]]');
            if ($nodesPSp->length > 0) {
                $nodePSp = $nodesPSp->item(0);
            }
            $positionSet = $position['placeholder']['name'];
        } else if (isset($position['placeholder']) && isset($position['placeholder']['descr'])) {
            // get placeholder by alt text
            if (!isset($position['placeholder']['position'])) {
                $nodesPSp = $slideXPath->query('//p:sp[.//p:cNvPr[@descr="'.$position['placeholder']['descr'].'"]]');
            } else if (isset($position['placeholder']['position'])) {
                // get placeholder by type and position
                $nodesPSp = $slideXPath->query('//p:sp[.//p:cNvPr[@descr="'.$position['placeholder']['descr'].'"]['.((int)$position['placeholder']['position']+1).']]');
            }
            if (isset($nodesPSp) && $nodesPSp->length > 0) {
                $nodePSp = $nodesPSp->item(0);
            }
            $positionSet = $position['placeholder']['descr'];
        } else if (isset($position['placeholder']) && isset($position['placeholder']['type'])) {
            if (!isset($position['placeholder']['position'])) {
                // get placeholder by type
                $nodesPSp = $slideXPath->query('//p:sp[.//p:ph[@type="'.$position['placeholder']['type'].'"]]');
            } else if (isset($position['placeholder']['position'])) {
                // get placeholder by type and position
                $nodesPSp = $slideXPath->query('//p:sp[.//p:ph[@type="'.$position['placeholder']['type'].'"]['.((int)$position['placeholder']['position']+1).']]');
            }
            if (isset($nodesPSp) && $nodesPSp->length > 0) {
                $nodePSp = $nodesPSp->item(0);
            }
            $positionSet = $position['placeholder']['type'];
        } else if (isset($position['placeholder']) && isset($position['placeholder']['position'])) {
            // get placeholder by position
            $nodesPSp = $slideXPath->query('//p:sp['.((int)$position['placeholder']['position']+1).']');
            if ($nodesPSp->length > 0) {
                $nodePSp = $nodesPSp->item(0);
            }
            $positionSet = $position['placeholder']['position'];
        }

        // create and add a new text box
        if (isset($position['new'])) {
            $textBox = new CreateTextBox();
            if (!isset($position['new']['textBoxStyles'])) {
                $position['new']['textBoxStyles'] = array();
            }
            $nodePSp = $textBox->addElementTextBox($slideDOM, $position['new'], $position['new']['textBoxStyles'], $options);
            $positionSet = 'new';
        }

        if (!isset($nodePSp)) {
            PhppptxLogger::logger('The chosen position \''.$positionSet.'\' is not valid. Use a valid position.', 'fatal');
        }

        return $nodePSp;
    }

    /**
     * Add external relationships
     *
     * @param array $relationships
     * @param string $activeSlideContentPath
     * @param DOMDocument $slideRelsDOM
     */
    protected function addExternalRelationships($relationships, $activeSlideContentPath, $slideRelsDOM)
    {
        foreach ($relationships as $relationship) {
            $this->handleExternalRelationship($slideRelsDOM, $relationship);
        }
    }

    /**
     * Handles external relationships from contents and PptxFragments
     *
     * @access protected
     * @param DOMDocument $slideRelsDOM
     * @param array $externalRelationship
     * @throws Exception slide position not valid
     */
    protected function handleExternalRelationship($slideRelsDOM, $externalRelationship)
    {
        if ($externalRelationship['type'] == 'hyperlink') {
            if (strpos($externalRelationship['hyperlink'], '#slide') === 0) {
                // slide position

                // get the slide path
                $slidePosition = str_replace('#slide', '', $externalRelationship['hyperlink']);
                $slidesContents = $this->zipPptx->getSlides();

                if (!isset($slidesContents[$slidePosition])) {
                    PhppptxLogger::logger('The slide position doesn\'t exist.', 'fatal');
                }

                // get the target without folders
                $hyperlinkTarget = str_replace('ppt/slides/', '', $slidesContents[$slidePosition]['path']);

                // add a rels
                $newRelationship = '<Relationship Id="rId'.$externalRelationship['id'].'" Target="'.$this->parseAndCleanTextString($hyperlinkTarget).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>';
                $this->generateRelationship($slideRelsDOM, $newRelationship);
            } else {
                // external

                // generate and add a rels
                $newRelationship = '<Relationship Id="rId'.$externalRelationship['id'].'" Target="'.$this->parseAndCleanTextString($externalRelationship['hyperlink']).'" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"/>';
                $this->generateRelationship($slideRelsDOM, $newRelationship);
            }
        }
    }

    /**
     * Inserts list values in a recursive way
     *
     * @access public
     * @param DOMDocument $slideDOM
     * @param DOMNode $nodeTxBody
     * @param array $listValues
     * @param array $listStyles
     * @param int $level
     */
    private function insertListValues($slideDOM, $activeSlideContent, $nodeTxBody, $listValues, $listStyles, $level = 0) {
        foreach ($listValues as $key => $value) {
            if (is_array($value) && !isset($value['text'])) {
                $this->insertListValues($slideDOM, $activeSlideContent, $nodeTxBody, $value, $listStyles, $level + 1);
            } else {
                // allow string as $contents instead of an array. Transform string to array
                if (!is_array($value) && !($value instanceof PptxFragment)) {
                    $contentsNormalized = array();
                    $contentsNormalized['text'] = $value;
                    $value = $contentsNormalized;
                }

                if (!($value instanceof PptxFragment)) {
                    // not PptxFragment. Create the text tags
                    $paragraphStyles = array();
                    $paragraphStyles['listLevel'] = $level;
                    $paragraphStyles['noBullet'] = false;
                    if (isset($listStyles[$level])) {
                        $paragraphStyles['listStyles'] = $listStyles[$level];
                    }

                    // if not using a subarray, generate it
                    if (isset($value['text'])) {
                        $value = array($value);
                    }

                    $text = new CreateText();
                    $newContentText = $text->createElementText($value, $paragraphStyles);
                } else {
                    // PptxFragment
                    $newContentText = (string)$value;

                    // insert the list level
                    $listContentDOM = $this->xmlUtilities->generateDomDocument($newContentText);
                    $nodesPpr = $listContentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'pPr');
                    if ($nodesPpr->length == 0) {
                        $nodesP = $listContentDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'p');
                        if ($nodesP->length > 0) {
                            $newPprFragment = $nodesP->item(0)->ownerDocument->createDocumentFragment();
                            $newPprFragment->appendXML('<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />');
                            $nodePpr = $nodesP->item(0)->insertBefore($newPprFragment, $nodesP->item(0)->firstChild);
                        }
                    } else {
                        $nodePpr = $nodesPpr->item(0);
                    }
                    if ($level > 0 && isset($nodePpr)) {
                        $nodePpr->setAttribute('lvl', $level);
                    }

                    $newContentText = $listContentDOM->saveXML($listContentDOM->documentElement);

                    // handle external relationships such as hyperlinks
                    $externalRelationships = $value->getExternalRelationships();
                    if (count($externalRelationships) > 0) {
                        $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $activeSlideContent['path']) . '.rels';
                        $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');
                        $this->addExternalRelationships($externalRelationships, $activeSlideContent['path'], $slideRelsDOM);

                        // refresh contents
                        $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

                        // free DOMDocument resources
                        $slideRelsDOM = null;
                    }
                }

                // append the new content
                $newTextFragment = $slideDOM->createDocumentFragment();
                $newTextFragment->appendXML($newContentText);
                $nodeTxBody->appendChild($newTextFragment);
            }
        }
    }

    /**
     * Parses and clean a text string to be added
     *
     * @access protected
     * @param string $content
     * @return string
     */
    protected function parseAndCleanTextString($content)
    {
        $content = $this->xmlUtilities->parseAndCleanTextString($content);

        return $content;
    }
}
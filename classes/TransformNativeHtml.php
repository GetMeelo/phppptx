<?php
/**
 * Transform PPTX to HTML using native PHP classes
 *
 * @category   Phppptx
 * @package    transform
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */

require_once __DIR__ . '/TransformPlugin.php';

class TransformNativeHtml
{
    /**
     *
     * @access protected
     * @var array
     */
    protected $css;

    /**
     *
     * @access protected
     * @var string
     */
    protected $currentSectionClassName;

    /**
     *
     * @access protected
     * @var string
     */
    protected $html;

    /**
     *
     * @access protected
     * @var TransformNativeHtmlPlugin
     */
    protected $htmlPlugin;

    /**
     *
     * @access protected
     * @var string
     */
    protected $javascript;

    /**
     *
     * @access protected
     * @var PptxStructure
     */
    protected $pptxStructure;

    /**
     *
     * @access protected
     * @var array
     */
    protected $themeInformation = array();

    /**
     *
     * @access protected
     * @var XmlUtilities XML Utilities classes
     */
    protected $xmlUtilities;

    /**
     * Constructor
     *
     * @access public
     * @param string|PptxStructure $source File path or PptxStructure to be transformed
     */
    public function __construct($source)
    {
        if ($source instanceof PptxStructure) {
            $this->pptxStructure = $source;
        } else {
            $this->pptxStructure = new PptxStructure();
            $this->pptxStructure->parsePptx($source);
        }

        $this->xmlUtilities = new XmlUtilities();
    }

    /**
     * Transform PPTX
     *
     * @param TransformNativeHtmlPlugin $htmlPlugin Plugin to be used to transform the contents
     * @param array $options
     *      'javaScriptAtTop' (bool) default as false. If true add JS in the head tag
     *      'returnHtmlStructure' (bool) if true return each element of the HTML using an array: css, javascript, metas, presentation. Default as false
     * @return string|array
     */
    public function transform(TransformNativeHtmlPlugin $htmlPlugin, $options = array())
    {
        $this->htmlPlugin = $htmlPlugin;
        $this->css['.presentation'] = '';

        $this->themeInformation = array(
            'color' => array(),
            'font' => array(
                'major' => array(),
                'minor' => array(),
            ),
        );
        $themeContents = $this->pptxStructure->getContentByType('themes');
        if (count($themeContents) > 0 && isset($themeContents[0]['content'])) {
            $this->addThemeStyles($themeContents[0]['content']);
        }

        // add default font
        if (isset($this->themeInformation['font']['major']['latin'])) {
            $this->css['.presentation'] .= 'font-family: "' . $this->themeInformation['font']['major']['latin'] . '";';
        }

        $tableStyles = $this->pptxStructure->getContentByType('tableStyles');
        if (count($tableStyles) > 0 && isset($tableStyles[0]['content'])) {
            $this->addTableStyles($tableStyles[0]['content']);
        }

        $this->html = '<div class="presentation">';

        $presentationContents = $this->pptxStructure->getContentByType('presentations');
        $slidesPresentation = $this->pptxStructure->getSlides();

        // get presentation common styles
        $presentationDom = $this->xmlUtilities->generateDomDocument($presentationContents[0]['content']);
        $presentationXPath = new DOMXPath($presentationDom);
        $presentationXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $presentationXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $nodesLvl1Pr = $presentationXPath->query('//p:defaultTextStyle/a:lvl1pPr');
        if ($nodesLvl1Pr->length > 0) {
            // paragraph styles
            $this->css['.presentation'] .= $this->getPprStyles($nodesLvl1Pr->item(0));
        }
        $nodesSldSz = $presentationDom->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sldSz');
        $this->css['.slide'] = '';
        if ($nodesSldSz->length > 0) {
            // width and height
            $this->css['.slide'] .= 'width:' . $this->htmlPlugin->transformSizes($nodesSldSz->item(0)->getAttribute('cx'), 'twips') . ';height:' . $this->htmlPlugin->transformSizes($nodesSldSz->item(0)->getAttribute('cy'), 'twips') . ';';
        }
        $this->css['.slide'] .= 'position: relative;';

        foreach ($slidesPresentation as $slide) {
            // get the slide contents
            $slideDom = $this->xmlUtilities->generateDomDocument($slide['content']);

            // avoid adding hidden slides
            $nodesSld = $slideDom->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sld');
            if ($nodesSld->length > 0) {
                if ($nodesSld->item(0)->hasAttribute('show') && $nodesSld->item(0)->getAttribute('show') == '0') {
                    continue;
                }
            }

            // transform the slide contents
            $this->html .= $this->transformSlide($slide);

            // free resources
            $slideDom = null;
        }

        $this->html .= '</div>';


        // get meta values
        $metaValues = $this->getMetaValues();

        // generate the CSS contents
        $cssContent = '';
        foreach ($this->css as $key => $value) {
            if (!empty($value)) {
                $cssContent .= $key . '{' . $value . '}';
            }
        }

        // clean CSS
        $cssContent = str_replace('##', '#', $cssContent);
        $cssContent = str_replace('._span', 'span', $cssContent);

        $presentationContent = $this->html;

        // add empty &nbps; to avoid empty paragraphs
        $presentationContent = str_replace('"></p>', '">&nbsp;</p>', $presentationContent);

        if (isset($options['returnHtmlStructure']) && $options['returnHtmlStructure']) {
            $output = array(
                'css' => $cssContent,
                'javascript' => $this->javascript,
                'metas' => $metaValues,
                'presentation' => $presentationContent,
            );
        } else {
            if (isset($options['javaScriptAtTop']) && $options['javaScriptAtTop']) {
                $output = $this->htmlPlugin->getBaseHTML() . '<head>' . $this->htmlPlugin->getBaseMeta() . $metaValues . $this->htmlPlugin->getBaseCSS() . $this->htmlPlugin->getBaseJavaScript() . $this->javascript . '<style>' . $cssContent . '</style></head><body>' . $presentationContent . '</body></html>';
            } else {
                $output = $this->htmlPlugin->getBaseHTML() . '<head>' . $this->htmlPlugin->getBaseMeta() . $metaValues . $this->htmlPlugin->getBaseCSS() . '<style>' . $cssContent . '</style></head><body>' . $presentationContent . $this->htmlPlugin->getBaseJavaScript() . $this->javascript . '</body></html>';
            }
        }

        return $output;
    }

    /**
     * Iterate the contents and transform them
     *
     * @param array $slide slide
     */
    public function transformSlide($slide)
    {
        $nodeClass = $this->htmlPlugin->generateClassName();
        $this->css['.' . $nodeClass] = '';

        // get slide DOMDocument and XPath elements
        $slideDom = $this->xmlUtilities->generateDomDocument($slide['content']);
        $slideLayoutDom = null;
        $slideMasterDom = null;
        $slideXPath = new DOMXPath($slideDom);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideRelsPath = str_replace('ppt/slides/', 'ppt/slides/_rels/', $slide['path']) . '.rels';
        $slideRelsDom = $this->pptxStructure->getContent($slideRelsPath, 'DOMDocument');
        $slideRelsXPath = new DOMXPath($slideRelsDom);
        $slideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $nodesSlideLayout = $slideRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"]');
        $slideLayoutRelsDom = null;
        if ($nodesSlideLayout->length > 0 && $nodesSlideLayout->item(0)->hasAttribute('Target')) {
            $slideLayoutPath = $nodesSlideLayout->item(0)->getAttribute('Target');
            $slideLayoutPath = str_replace('../', 'ppt/', $slideLayoutPath);
            $slideLayoutDom = $this->pptxStructure->getContent($slideLayoutPath, 'DOMDocument');

            // get slide master
            $slideLayoutRelsPath = str_replace('ppt/slideLayouts/', 'ppt/slideLayouts/_rels/', $slideLayoutPath) . '.rels';
            $slideLayoutRelsDom = $this->pptxStructure->getContent($slideLayoutRelsPath, 'DOMDocument');
            $slideLayoutRelsXPath = new DOMXPath($slideLayoutRelsDom);
            $slideLayoutRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
            $nodesSlideMaster = $slideLayoutRelsXPath->query('//xmlns:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"]');
            if ($nodesSlideMaster->length > 0 && $nodesSlideMaster->item(0)->hasAttribute('Target')) {
                $slideMasterPath = $nodesSlideMaster->item(0)->getAttribute('Target');
                $slideMasterPath = str_replace('../', 'ppt/', $slideMasterPath);
                $slideMasterDom = $this->pptxStructure->getContent($slideMasterPath, 'DOMDocument');

                // map theme information merging the master map styles
                $nodesClrMap = $slideMasterDom->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'clrMap');
                if ($nodesClrMap->length > 0) {
                    foreach ($nodesClrMap->item(0)->attributes as $attributeMap) {
                        if (isset($this->themeInformation['color'][$attributeMap->name]) && isset($this->themeInformation['color'][$attributeMap->value])) {
                            $this->themeInformation['color'][$attributeMap->name] = $this->themeInformation['color'][$attributeMap->value];
                        } else if (isset($this->themeInformation['color'][$attributeMap->value])) {
                            $this->themeInformation['color'][$attributeMap->name] = $this->themeInformation['color'][$attributeMap->value];
                        }
                    }
                }
            }
        }

        // keep path to the slide as id to be used as anchor links
        $html = '<div id="' . str_replace(array('ppt/slides/', '.'), array('', '_'), $slide['path']) . '" class="slide ' . ' ' . $nodeClass . '">';

        // slide background
        $nodesBg = $slideDom->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'bg');
        if ($nodesBg->length > 0) {
            // background color
            $nodesSolidFill = $nodesBg->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'solidFill');
            if ($nodesSolidFill->length > 0) {
                $this->css['.' . $nodeClass] = 'background-color: #' . $this->getColor($nodesSolidFill->item(0)) . ';';
            }

            // background image
            $nodesBlipFill = $nodesBg->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blipFill');
            if ($nodesBlipFill->length > 0) {
                $nodesBlip = $nodesBlipFill->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip');
                if ($nodesBlip->length > 0 && $nodesBlip->item(0)->hasAttribute('r:embed')) {
                    $nodeRelationship = $slideRelsXPath->query('.//xmlns:Relationship[@Id="'.$nodesBlip->item(0)->getAttribute('r:embed').'"]');
                    $target = str_replace('../', 'ppt/', $nodeRelationship->item(0)->getAttribute('Target'));
                    $fileInfo = pathinfo($target);
                    $ext = $fileInfo['extension'];
                    $fileString = $this->pptxStructure->getContent($target);

                    if ($this->htmlPlugin->getFilesAsBase64()) {
                        $src = 'data:image/' . $ext . ';base64,' . base64_encode($fileString);
                    } else {
                        $src = $this->htmlPlugin->getOutputFilesPath() . $fileInfo['filename'] . '.' . $ext;
                        file_put_contents($src, $fileString);
                    }

                    $this->css['.' . $nodeClass] .= 'background-image: url(' . $src . ');';

                    // check if the background image must cover the entire slide or it's repeated
                    $nodesStrech = $nodesBlipFill->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'stretch');
                    if ($nodesStrech->length > 0) {
                        $nodesFillRect = $nodesStrech->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'fillRect');
                        if ($nodesFillRect->length > 0) {
                            $this->css['.' . $nodeClass] .= 'background-size: contain; background-repeat: no-repeat; background-position: center center;background-size:100% 100%;';
                        }
                    }

                    // alpha transparency
                    $nodesAlphaModFix = $nodesBlipFill->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'alphaModFix');
                    if ($nodesAlphaModFix->length > 0 && $nodesAlphaModFix->item(0)->hasAttribute('amt')) {
                        $this->css['.' . $nodeClass] .= 'background-color: rgba(255, 255, 255, ' . (1 - (int)$nodesAlphaModFix->item(0)->getAttribute('amt') / 100000) . ');background-blend-mode: overlay;';
                    }
                }
            }
        }

        // slide tree
        $nodesSpTree = $slideDom->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'spTree');
        if ($nodesSpTree->length > 0) {
            foreach ($nodesSpTree->item(0)->childNodes as $childNode) {
                // open tag
                switch ($childNode->nodeName) {
                    case 'p:graphicFrame':
                        $html .= $this->transformP_GRAPHICFRAME($childNode, $slideDom, $slideRelsDom, $slideLayoutDom, $slideLayoutRelsDom, $slideMasterDom);
                        break;
                    case 'p:pic':
                        $html .= $this->transformP_PIC($childNode, $slideDom, $slideRelsDom, $slideLayoutDom, $slideLayoutRelsDom, $slideMasterDom);
                        break;
                    case 'p:sp':
                        $html .= $this->transformP_SP($childNode, $slideDom, $slideRelsDom, $slideLayoutDom, $slideLayoutRelsDom, $slideMasterDom);
                        break;
                }
            }
        }
        $html .= '</div>';

        return $html;
    }

    /**
     * Add table styles
     *
     * @param string $tableStyles Table styles
     * @return string CSS styles
     */
    protected function addTableStyles($tableStyles)
    {
        $css = '';

        $tableStylesDom = $this->xmlUtilities->generateDomDocument($tableStyles);

        $nodesTblStyle = $tableStylesDom->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'tblStyle');
        foreach ($nodesTblStyle as $nodeTblStyle) {
            if ($nodeTblStyle->hasAttribute('styleId')) {
                $styleClass = 't' . str_replace(array('{', '}'), '', $nodeTblStyle->getAttribute('styleId'));
                $tablePositions = array(
                    'wholeTbl' => 'table.' . $styleClass . ' td',
                    'band1H' => 'table.' . $styleClass . ' tr:nth-child(even) td',
                    'band2H' => 'table.' . $styleClass . ' tr:nth-child(odd) td',
                    'band1V' => 'table.' . $styleClass . ' tr td:nth-child(odd)',
                    'band2V' => 'table.' . $styleClass . ' tr td:nth-child(even)',
                    'lastCol' => 'table.' . $styleClass . ' tr td:last-child',
                    'firstCol' => 'table.' . $styleClass . ' tr td:first-child',
                    'lastRow' => 'table.' . $styleClass . ' tr:last-child td',
                    'firstRow' => 'table.' . $styleClass . ' tr:first-child td',
                );

                foreach ($tablePositions as $position => $tableCssName) {
                    $tableCssValue = '';
                    $nodesTblStyleElement = $nodeTblStyle->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', $position);
                    if ($nodesTblStyleElement->length > 0) {
                        $tableCssValue = $this->getTableStyles($nodesTblStyleElement->item(0));
                        $this->css[$tableCssName] = $tableCssValue;
                    }
                }
            }
        }

        return $css;
    }

    /**
     * Theme file
     *
     * @param string $theme Theme content
     */
    protected function addThemeStyles($theme)
    {
        $themeDom = $this->xmlUtilities->generateDomDocument($theme);

        // colors
        $elementsClrScheme = $themeDom->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'clrScheme');
        if ($elementsClrScheme->length > 0) {
            foreach ($elementsClrScheme->item(0)->childNodes as $childNode) {
                $nodeName = str_replace('a:', '', $childNode->nodeName);
                if ($childNode->firstChild && $childNode->firstChild->nodeName == 'a:srgbClr' && $childNode->firstChild->hasAttribute('val')) {
                    $this->themeInformation['color'][$nodeName] = $childNode->firstChild->getAttribute('val');
                } elseif ($childNode->firstChild && $childNode->firstChild->nodeName == 'a:sysClr' && $childNode->firstChild->hasAttribute('lastClr')) {
                    $this->themeInformation['color'][$nodeName] = $childNode->firstChild->getAttribute('lastClr');
                }
            }
        }

        // fonts
        $nodesMajorFont = $themeDom->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'majorFont');
        if ($nodesMajorFont->length > 0) {
            $nodesMajorFontLatin = $nodesMajorFont->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'latin');
            if ($nodesMajorFontLatin->length > 0 && $nodesMajorFontLatin->item(0)->hasAttribute('typeface') && !empty($nodesMajorFontLatin->item(0)->getAttribute('typeface'))) {
                $this->themeInformation['font']['major']['latin'] = $nodesMajorFontLatin->item(0)->getAttribute('typeface');
            }
        }
        $nodesMinorFont = $themeDom->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'minorFont');
        if ($nodesMinorFont->length > 0) {
            $nodesMinorFontLatin = $nodesMinorFont->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'latin');
            if ($nodesMinorFontLatin->length > 0 && $nodesMinorFontLatin->item(0)->hasAttribute('typeface') && !empty($nodesMinorFontLatin->item(0)->getAttribute('typeface'))) {
                $this->themeInformation['font']['minor']['latin'] = $nodesMinorFontLatin->item(0)->getAttribute('typeface');
            }
        }
    }

    /**
     * Normalize border styles
     *
     * @param string $style Border style
     * @return string Styles
     */
    protected function getBorderStyle($style)
    {
        $borderStyle = 'solid';
        switch ($style) {
            case 'dashed':
                $borderStyle ='dashed';
                break;
            case 'dotted':
                $borderStyle ='dotted';
                break;
            case 'double':
                $borderStyle ='double';
                break;
            case 'nil':
            case 'none':
                $borderStyle = 'none';
                break;
            case 'single':
                $borderStyle = 'solid';
                break;
            default:
                $borderStyle = 'solid';
                break;
        }

        return $borderStyle;
    }

    /**
     * Get color
     *
     * @param DOMNode $nodeColor
     * @return string
     */
    protected function getColor($nodeColor)
    {
        $color = '';

        // scheme color
        $nodesSchemeClr = $nodeColor->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'schemeClr');
        if ($nodesSchemeClr->length > 0 && $nodesSchemeClr->item(0)->hasAttribute('val') && isset($this->themeInformation['color'][$nodesSchemeClr->item(0)->getAttribute('val')])) {
            $color .= $this->themeInformation['color'][$nodesSchemeClr->item(0)->getAttribute('val')];
        }
        // RGB color
        $nodesSrgbClr = $nodeColor->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'srgbClr');
        if ($nodesSrgbClr->length > 0 && $nodesSrgbClr->item(0)->hasAttribute('val')) {
            $color .= $nodesSrgbClr->item(0)->getAttribute('val');
        }

        return $color;
    }

    /**
     * Meta values
     *
     * @return string metas
     */
    protected function getMetaValues()
    {
        $documentCoreContent = $this->pptxStructure->getContent('docProps/core.xml');

        $tags = '';

        if ($documentCoreContent) {
            $xmlCoreContent = $this->xmlUtilities->generateDomDocument($documentCoreContent);
            foreach ($xmlCoreContent->childNodes->item(0)->childNodes as $prop) {
                switch ($prop->tagName) {
                    case 'dc:title':
                        $tags .= '<title>' . $prop->nodeValue . '</title>';
                        break;
                    case 'dc:creator':
                        $tags .= '<meta name="author" content="' . $prop->nodeValue . '">';
                        break;
                    case 'cp:keywords':
                        $tags .= '<meta name="keywords" content="' . $prop->nodeValue . '">';
                        break;
                    case 'dc:description':
                        $tags .= '<meta name="description" content="' . $prop->nodeValue . '">';
                        break;
                    default:
                        break;
                }
            }
        }

        return $tags;
    }

    /**
     * Get bodyPr styles
     *
     * @param DOMNode $nodeBodyPr
     * @return string
     */
    protected function getBodyPrStyles($nodeBodyPr)
    {
        $css = '';

        if ($nodeBodyPr->hasAttribute('anchor')) {
            switch ($nodeBodyPr->getAttribute('anchor')) {
                case 'b':
                    $css .= 'justify-content: end;';
                    break;
                case 'ctr':
                    $css .= 'justify-content: center;';
                    break;
                case 't':
                    $css .= 'justify-content: start;';
                    break;
                default:
                    break;
            }
        }

        return $css;
    }

    /**
     * Get bodyPr p styles
     *
     * @param DOMNode $nodeBodyPr
     * @return string
     */
    protected function getBodyPrStylesP($nodeBodyPr)
    {
        $css = '';

        if ($nodeBodyPr->hasAttribute('vert')) {
            switch ($nodeBodyPr->getAttribute('vert')) {
                case 'vert':
                    $css .= 'transform: rotate(90deg);';
                    break;
                case 'vert270':
                    $css .= 'transform: rotate(270deg);';
                    break;
                default:
                    break;
            }
        }

        return $css;
    }

    /**
     * Get spPr styles
     *
     * @param DOMNode $nodeSpPr
     * @return string
     */
    protected function getSpPrStyles($nodeSpPr)
    {
        $css = '';

        // parse children
        if ($nodeSpPr->hasChildNodes()) {
            foreach ($nodeSpPr->childNodes as $childNode) {
                switch ($childNode->nodeName) {
                    case 'a:solidFill':
                        $nodesSrgbClr = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'srgbClr');
                        if ($nodesSrgbClr->length > 0 && $nodesSrgbClr->item(0)->hasAttribute('val')) {
                            $css .= 'background-color: #' . $nodesSrgbClr->item(0)->getAttribute('val') . ';';
                        }
                        break;
                    case 'a:ln':
                        $nodesSolidFill = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'solidFill');
                        $css .= 'border-style: solid;';
                        if ($nodesSolidFill->length > 0) {
                            $css .= 'border-color: #' . $this->getColor($nodesSolidFill->item(0)) . ';';
                        }
                        if ($childNode->hasAttribute('w')) {
                            $css .= 'border-width: ' . $this->htmlPlugin->transformSizes($childNode->getAttribute('w'), 'twips') . ';';
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        return $css;
    }

    /**
     * Get pPr styles
     *
     * @param DOMNode $nodePpr
     * @return string
     */
    protected function getPprStyles($nodePpr)
    {
        $css = '';

        // parse attributes
        foreach ($nodePpr->attributes as $attribute) {
            switch ($attribute->name) {
                case 'algn':
                    switch ($attribute->value) {
                        case 'ctr':
                            $css .= 'align-self: center;text-align: center;';
                            break;
                        case 'l':
                            $css .= 'align-self: start;text-align: left;';
                            break;
                        case 'r':
                            $css .= 'align-self: end;text-align: right;';
                            break;
                        default:
                            break;
                    }
                    break;
                case 'marL':
                    if ($attribute->value > 0) {
                        $css .= 'margin-left: ' . $this->htmlPlugin->transformSizes($attribute->value, 'twips') . ';';
                    }
                    break;
                default:
                    break;
            }
        }

        // parse children
        if ($nodePpr->hasChildNodes()) {
            foreach ($nodePpr->childNodes as $childNode) {
                switch ($childNode->nodeName) {
                    case 'a:defRPr':
                        if ($childNode->hasAttribute('sz')) {
                            $css .= 'font-size:' . $this->htmlPlugin->transformSizes($childNode->getAttribute('sz'), 'hundreds', null, false) . ';';
                        }
                        break;
                    case 'a:spcBef':
                        $nodesSpcPts = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'spcPts');
                        if ($nodesSpcPts->length > 0 && $nodesSpcPts->item(0)->hasAttribute('val')) {
                            $css .= 'margin-top: ' . $this->htmlPlugin->transformSizes($nodesSpcPts->item(0)->getAttribute('val'), 'hundreds', null, false) . ';';
                        }
                        break;
                    case 'a:spcAft':
                        $nodesSpcPts = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'spcPts');
                        if ($nodesSpcPts->length > 0 && $nodesSpcPts->item(0)->hasAttribute('val')) {
                            $css .= 'margin-bottom: ' . $this->htmlPlugin->transformSizes($nodesSpcPts->item(0)->getAttribute('val'), 'hundreds', null, false) . ';';
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        return $css;
    }

    /**
     * Get rPr styles
     *
     * @param DOMNode $nodeRpr
     * @return string
     */
    protected function getRprStyles($nodeRpr)
    {
        $css = '';

        // parse attributes
        foreach ($nodeRpr->attributes as $attribute) {
            switch ($attribute->name) {
                case 'b':
                    if ($attribute->value == '1') {
                        $css .= 'font-weight: bold;';
                    }
                    break;
                case 'i':
                    if ($attribute->value == '1') {
                        $css .= 'font-style: italic;';
                    }
                    break;
                case 'strike':
                    if ($attribute->value == 'sngStrike') {
                        $css .= 'text-decoration: line-through;';
                    }
                    break;
                case 'sz':
                    if ($attribute->value > 0) {
                        $css .= 'font-size:' . $this->htmlPlugin->transformSizes($attribute->value, 'hundreds', null, false) . ';';
                    }
                    break;
                case 'u':
                    if ($attribute->value == 'sng') {
                        $css .= 'text-decoration: underline;';
                    }
                    break;
                default:
                    break;
            }
        }

        // parse children
        if ($nodeRpr->hasChildNodes()) {
            foreach ($nodeRpr->childNodes as $childNode) {
                switch ($childNode->nodeName) {
                    case 'a:highlight':
                        $nodesSrgbClr = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'srgbClr');
                        if ($nodesSrgbClr->length > 0 && $nodesSrgbClr->item(0)->hasAttribute('val')) {
                            $css .= 'background-color: #' . $nodesSrgbClr->item(0)->getAttribute('val') . ';';
                        }
                        break;
                    case 'a:solidFill':
                        $nodesSrgbClr = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'srgbClr');
                        if ($nodesSrgbClr->length > 0 && $nodesSrgbClr->item(0)->hasAttribute('val')) {
                            $css .= 'color: #' . $nodesSrgbClr->item(0)->getAttribute('val') . ';';
                        }
                        $nodesSchemeClr = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'schemeClr');
                        if ($nodesSchemeClr->length > 0 && $nodesSchemeClr->item(0)->hasAttribute('val') && isset($this->themeInformation['color'][$nodesSchemeClr->item(0)->getAttribute('val')])) {
                            $css .= 'color: #' . $this->themeInformation['color'][$nodesSchemeClr->item(0)->getAttribute('val')] . ';';
                        }
                        break;
                    case 'a:latin':
                        if ($childNode->hasAttribute('typeface')) {
                            $css .= 'font-family: "' . $childNode->getAttribute('typeface') . '";';
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        return $css;
    }

    /**
     * Get shape details: position and size
     *
     * @param DOMNode $childNode
     * @param DOMXPath $slideXPath
     * @param DOMDocument $slideLayoutDom
     * @param DOMDocument $slideMasterDom
     * @return array
     */
    protected function getShapeDetails($childNode, $slideXPath, $slideLayoutDom, $slideMasterDom)
    {
        // get shape idx, id and type
        $nodesCNvPr = $slideXPath->query('.//p:nvSpPr/p:cNvPr', $childNode);
        $shapePlaceholder = array();
        if ($nodesCNvPr->length > 0 && $nodesCNvPr->item(0)->hasAttribute('id')) {
            $shapePlaceholder['id'] = $nodesCNvPr->item(0)->getAttribute('id');
        }
        $nodesPh = $slideXPath->query('.//p:nvPr/p:ph', $childNode);
        if ($nodesPh->length > 0) {
            if ($nodesPh->item(0)->hasAttribute('idx')) {
                $shapePlaceholder['idx'] = $nodesPh->item(0)->getAttribute('idx');
            }
            if ($nodesPh->item(0)->hasAttribute('type')) {
                $shapePlaceholder['type'] = $nodesPh->item(0)->getAttribute('type');
            }
        }

        // check if the p:psp has position and size
        $nodesXfrm = $slideXPath->query('.//p:spPr/a:xfrm | .//p:xfrm', $childNode);

        // there's no position and size, get them from the slide layout
        $slideLayoutXPath = new DOMXPath($slideLayoutDom);
        $slideLayoutXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideLayoutXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        if ($nodesXfrm->length == 0 && !is_null($slideLayoutDom) && (isset($shapePlaceholder['idx']) || isset($shapePlaceholder['type']))) {
            // idx and type
            if ((isset($shapePlaceholder['idx']) && isset($shapePlaceholder['type']))) {
                // idx and type
                $nodesXfrm = $slideLayoutXPath->query('//p:sp[.//p:nvPr/p:ph[@idx="' . $shapePlaceholder['idx'] . '" and @type="' . $shapePlaceholder['type'] . '"]]/p:spPr/a:xfrm');
            } elseif (isset($shapePlaceholder['type'])) {
                // type
                $nodesXfrm = $slideLayoutXPath->query('//p:sp[.//p:nvPr/p:ph[@type="' . $shapePlaceholder['type'] . '"]]/p:spPr/a:xfrm');
            } elseif (isset($shapePlaceholder['idx'])) {
                // idx
                $nodesXfrm = $slideLayoutXPath->query('//p:sp[.//p:nvPr/p:ph[@idx="' . $shapePlaceholder['idx'] . '"]]/p:spPr/a:xfrm');
            }
        }

        // there's no position and size, get them from the slide master
        $slideMasterXPath = new DOMXPath($slideMasterDom);
        $slideMasterXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideMasterXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        if ($nodesXfrm->length == 0 && !is_null($slideMasterDom) && (isset($shapePlaceholder['idx']) || isset($shapePlaceholder['type']))) {
            if ((isset($shapePlaceholder['idx']) && isset($shapePlaceholder['type']))) {
                // idx and type
                $nodesXfrm = $slideMasterXPath->query('//p:sp[.//p:nvPr/p:ph[@idx="' . $shapePlaceholder['idx'] . '" and @type="' . $shapePlaceholder['type'] . '"]]/p:spPr/a:xfrm');
            } elseif (isset($shapePlaceholder['type'])) {
                // type
                $nodesXfrm = $slideMasterXPath->query('//p:sp[.//p:nvPr/p:ph[@type="' . $shapePlaceholder['type'] . '"]]/p:spPr/a:xfrm');
            } elseif (isset($shapePlaceholder['idx'])) {
                // idx
                $nodesXfrm = $slideMasterXPath->query('//p:sp[.//p:nvPr/p:ph[@idx="' . $shapePlaceholder['idx'] . '"]]/p:spPr/a:xfrm');
            }
        }

        $shapeDrawing = array();

        if ($nodesXfrm->length > 0) {
            $shapePosition = $nodesXfrm->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'off');
            // position
            if ($shapePosition->length > 0) {
                $shapeDrawing['x'] = $this->htmlPlugin->transformSizes($shapePosition->item(0)->getAttribute('x'), 'twips');
                $shapeDrawing['y'] = $this->htmlPlugin->transformSizes($shapePosition->item(0)->getAttribute('y'), 'twips');
            }

            // size
            $shapeSize = $nodesXfrm->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'ext');
            if ($shapeSize->length > 0) {
                $shapeDrawing['width'] = $this->htmlPlugin->transformSizes($shapeSize->item(0)->getAttribute('cx'), 'twips');
                $shapeDrawing['height'] = $this->htmlPlugin->transformSizes($shapeSize->item(0)->getAttribute('cy'), 'twips');
            }

            // rotation
            if ($nodesXfrm->item(0)->hasAttribute('rot')) {
                $shapeDrawing['rotation'] = $this->htmlPlugin->transformSizes($nodesXfrm->item(0)->getAttribute('rot'), 'twips', 'deg', false);
            }

            // border
            $shapeLine = $nodesXfrm->item(0)->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'ln');
            if ($shapeLine->length > 0 && $shapeLine->item(0)->hasAttribute('w')) {
                $shapeDrawing['border'] = array();
                $shapeDrawing['border']['size'] = $this->htmlPlugin->transformSizes($shapeLine->item(0)->getAttribute('w'), 'twips');
                $shapeDrawing['border']['color'] = '#000000';
                // scheme color
                $nodesSchemeClr = $shapeLine->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'schemeClr');
                if ($nodesSchemeClr->length > 0 && $nodesSchemeClr->item(0)->hasAttribute('val') && isset($this->themeInformation['color'][$nodesSchemeClr->item(0)->getAttribute('val')])) {
                    $shapeDrawing['border']['color'] = '#' . $this->themeInformation['color'][$nodesSchemeClr->item(0)->getAttribute('val')] . ';';
                }
                // RGB color
                $nodesSrgbClr = $shapeLine->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'srgbClr');
                if ($nodesSrgbClr->length > 0 && $nodesSrgbClr->item(0)->hasAttribute('val')) {
                    $shapeDrawing['border']['color'] = '#' . $nodesSrgbClr->item(0)->getAttribute('val') . ';';
                }
            }

            // background color
            $nodesSolidFill = $nodesXfrm->item(0)->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'solidFill');
            foreach ($nodesSolidFill as $nodeSolidFill) {
                if ($nodeSolidFill->parentNode->tagName == 'p:spPr') {
                    $shapeDrawing['backgroundColor'] = '#' . $this->getColor($nodeSolidFill) . ';';
                }
            }
        }

        return array($shapeDrawing, $shapePlaceholder, $slideLayoutXPath, $slideMasterXPath);
    }

    /**
     * Get table styles
     *
     * @param DOMNode $nodeTbl
     * @return string CSS styles
     */
    protected function getTableStyles($nodeTbl)
    {
        $css = '';

        $nodesTcStyle = $nodeTbl->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'tcStyle');
        if ($nodesTcStyle->length > 0) {
            $nodesFill = $nodesTcStyle->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'solidFill');
            foreach ($nodesFill as $nodeFill) {
                if ($nodeFill->parentNode->tagName == 'a:fill') {
                    $css .= 'background-color: #' . $this->getColor($nodeFill) . ';';
                }
            }
        }

        return $css;
    }

    /**
     * Set CSS styles from slide master
     *
     * @param DOMDocument $dom
     * @param DOMXPath $xpath
     * @param string $shapeType
     * @param string $nodeClass
     * @param string $elementTagExtra
     */
    protected function setCssStylesMaster($dom, $xpath, $shapeType, $nodeClass, $elementTagExtra)
    {
        $styleTypeMap = [
            'title' => 'titleStyle',
            'body' => 'bodyStyle',
        ];

        if (isset($styleTypeMap[$shapeType])) {
            $nodesLvl1Pr = $xpath->query('//p:' . $styleTypeMap[$shapeType] . '/a:lvl1pPr');
            if ($nodesLvl1Pr->length > 0) {
                // paragraph styles
                $this->css['.' . $nodeClass . ' ' . $elementTagExtra] .= $this->getPprStyles($nodesLvl1Pr->item(0));
            }
        }
    }

    /**
     * Set CSS styles for txBody
     *
     * @param DOMDocument $dom
     * @param DOMXPath $xpath
     * @param array $shapePlaceholder
     * @param string $nodeClass
     * @param string $elementTagExtra
     * @param array $types id, idx, type
     */
    protected function setCssStylesTxBody($dom, $xpath, $shapePlaceholder, $nodeClass, $elementTagExtra, $types)
    {
        if (in_array('id', $types) && isset($shapePlaceholder['id'])) {
            // id
            $nodesLvl1Pr = $xpath->query('//p:sp[.//p:cNvPr[@id="' . $shapePlaceholder['id'] . '"]]/p:txBody/a:lstStyle/a:lvl1pPr');
            if ($nodesLvl1Pr->length > 0) {
                // paragraph styles
                $this->css['.' . $nodeClass . ' ' . $elementTagExtra] .= $this->getPprStyles($nodesLvl1Pr->item(0));
            }
            $nodesBodyPr = $xpath->query('//p:sp[.//p:cNvPr[@id="' . $shapePlaceholder['id'] . '"]]/p:txBody/a:bodyPr');
            if ($nodesBodyPr->length > 0) {
                // container styles
                $this->css['.' . $nodeClass] .= $this->getBodyPrStyles($nodesBodyPr->item(0));
                $this->css['.' . $nodeClass . ' ' . $elementTagExtra] .= $this->getBodyPrStylesP($nodesBodyPr->item(0));
            }
        } elseif (in_array('idx', $types) || in_array('type', $types)) {
            if (in_array('idx', $types) && in_array('type', $types) && isset($shapePlaceholder['idx']) && isset($shapePlaceholder['type'])) {
                // idx and type
                $nodesLvl1Pr = $xpath->query('//p:sp[.//p:nvPr/p:ph[@idx="' . $shapePlaceholder['idx'] . '" and @type="' . $shapePlaceholder['type'] . '"]]/p:txBody/a:lstStyle/a:lvl1pPr');
                if ($nodesLvl1Pr->length > 0) {
                    // paragraph styles
                    $this->css['.' . $nodeClass . ' ' . $elementTagExtra] .= $this->getPprStyles($nodesLvl1Pr->item(0));
                }
                $nodesBodyPr = $xpath->query('//p:sp[.//p:nvPr/p:ph[@idx="' . $shapePlaceholder['idx'] . '" and @type="' . $shapePlaceholder['type'] . '"]]/p:txBody/a:bodyPr');
                if ($nodesBodyPr->length > 0) {
                    // container styles
                    $this->css['.' . $nodeClass] .= $this->getBodyPrStyles($nodesBodyPr->item(0));
                    $this->css['.' . $nodeClass . ' ' . $elementTagExtra] .= $this->getBodyPrStylesP($nodesBodyPr->item(0));
                }
            } elseif (in_array('type', $types) && isset($shapePlaceholder['type'])) {
                // type
                $nodesLvl1Pr = $xpath->query('//p:sp[.//p:nvPr/p:ph[@type="' . $shapePlaceholder['type'] . '"]]/p:txBody/a:lstStyle/a:lvl1pPr');
                if ($nodesLvl1Pr->length > 0) {
                    // paragraph styles
                    $this->css['.' . $nodeClass . ' ' . $elementTagExtra] .= $this->getPprStyles($nodesLvl1Pr->item(0));
                }
                $nodesBodyPr = $xpath->query('//p:sp[.//p:nvPr/p:ph[@type="' . $shapePlaceholder['type'] . '"]]/p:txBody/a:bodyPr');
                if ($nodesBodyPr->length > 0) {
                    // container styles
                    $this->css['.' . $nodeClass] .= $this->getBodyPrStyles($nodesBodyPr->item(0));
                    $this->css['.' . $nodeClass . ' ' . $elementTagExtra] .= $this->getBodyPrStylesP($nodesBodyPr->item(0));
                }
            } elseif (in_array('idx', $types) && isset($shapePlaceholder['idx'])) {
                // idx
                $nodesLvl1Pr = $xpath->query('//p:sp[.//p:nvPr/p:ph[@idx="' . $shapePlaceholder['idx'] . '"]]/p:txBody/a:lstStyle/a:lvl1pPr');
                if ($nodesLvl1Pr->length > 0) {
                    // paragraph styles
                    $this->css['.' . $nodeClass . ' ' . $elementTagExtra] .= $this->getPprStyles($nodesLvl1Pr->item(0));
                }
                $nodesBodyPr = $xpath->query('//p:sp[.//p:nvPr/p:ph[@idx="' . $shapePlaceholder['idx'] . '"]]/p:txBody/a:bodyPr');
                if ($nodesBodyPr->length > 0) {
                    // container styles
                    $this->css['.' . $nodeClass] .= $this->getBodyPrStyles($nodesBodyPr->item(0));
                }
            }
        }
    }

    /**
     * Transform p:sp tag
     *
     * @param DOMNode $childNode
     * @param DOMDocument $slideDom
     * @param DOMDocument $slideRelsDom
     * @param DOMDocument $slideLayoutDom
     * @param DOMDocument $slideLayoutRelsDom
     * @param DOMDocument $slideMasterDom
     */
    protected function transformP_SP($childNode, $slideDom, $slideRelsDom, $slideLayoutDom = null, $slideLayoutRelsDom = null, $slideMasterDom = null)
    {
        $nodeClass = $this->htmlPlugin->generateClassName();
        $this->css['.' . $nodeClass] = '';
        $html = '';

        // default element
        $elementTag = $this->htmlPlugin->getTag('shape');

        $slideXPath = new DOMXPath($slideDom);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

        $slideRelsXPath = new DOMXPath($slideRelsDom);
        $slideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // get shape information
        list($shapeDrawing, $shapePlaceholder, $slideLayoutXPath, $slideMasterXPath) = $this->getShapeDetails($childNode, $slideXPath, $slideLayoutDom, $slideMasterDom);
        $this->css['.' . $nodeClass] .= 'display: flex;flex-direction: column;position: absolute;top:' . $shapeDrawing['y'] . ';left:' . $shapeDrawing['x'] . ';width:' . $shapeDrawing['width'] . ';height:' . $shapeDrawing['height'] . ';';
        if (isset($shapeDrawing['rotation'])) {
            $this->css['.' . $nodeClass] .= 'transform: rotate(' . $shapeDrawing['rotation'] . ');';
        }
        if (isset($shapeDrawing['backgroundColor'])) {
            $this->css['.' . $nodeClass] .= 'background-color: ' . $shapeDrawing['backgroundColor'] . ';';
        }
        if (isset($shapeDrawing['border'])) {
            $this->css['.' . $nodeClass] .= 'border: ' . $shapeDrawing['border']['size'] . ' solid ' . $shapeDrawing['border']['color'] . ';';
        }

        $nodesSpPr = $slideXPath->query('.//p:spPr', $childNode);
        if ($nodesSpPr->length > 0 && $nodesSpPr->item(0)->hasChildNodes()) {
            $this->css['.' . $nodeClass] .= $this->getSpPrStyles($nodesSpPr->item(0));
        }

        $html .= '<'.$elementTag.' class="'.$nodeClass.' ' . $this->htmlPlugin->getExtraClass('shape') . '">';

        // background image
        $nodesBlipFill = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blipFill');
        if ($nodesBlipFill->length > 0) {
            $nodesBlip = $nodesBlipFill->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip');
            if ($nodesBlip->length > 0 && $nodesBlip->item(0)->hasAttribute('r:embed')) {
                $nodeRelationship = $slideRelsXPath->query('.//xmlns:Relationship[@Id="'.$nodesBlip->item(0)->getAttribute('r:embed').'"]');
                $target = str_replace('../', 'ppt/', $nodeRelationship->item(0)->getAttribute('Target'));
                $fileInfo = pathinfo($target);
                $ext = $fileInfo['extension'];
                $fileString = $this->pptxStructure->getContent($target);

                if ($this->htmlPlugin->getFilesAsBase64()) {
                    $src = 'data:image/' . $ext . ';base64,' . base64_encode($fileString);
                } else {
                    $src = $this->htmlPlugin->getOutputFilesPath() . $fileInfo['filename'] . '.' . $ext;
                    file_put_contents($src, $fileString);
                }

                $this->css['.' . $nodeClass] .= 'background-image: url(' . $src . ');';

                // check if the background image must cover the entire slide or it's repeated
                $nodesStrech = $nodesBlipFill->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'stretch');
                if ($nodesStrech->length > 0) {
                    $nodesFillRect = $nodesStrech->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'fillRect');
                    if ($nodesFillRect->length > 0) {
                        $this->css['.' . $nodeClass] .= 'background-size: contain; background-repeat: no-repeat; background-position: center center;background-size:100% 100%;';
                    }
                }

                // alpha transparency
                $nodesAlphaModFix = $nodesBlipFill->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'alphaModFix');
                if ($nodesAlphaModFix->length > 0 && $nodesAlphaModFix->item(0)->hasAttribute('amt')) {
                    $this->css['.' . $nodeClass] .= 'background-color: rgba(255, 255, 255, ' . (1 - (int)$nodesAlphaModFix->item(0)->getAttribute('amt') / 100000) . ');background-blend-mode: overlay;';
                }
            }
        }

        $html .= $this->transformP_TXBODY($childNode, $nodeClass, $slideDom, $slideXPath, $slideRelsDom, $slideRelsXPath, $slideLayoutDom, $slideLayoutXPath, $slideMasterDom, $slideMasterXPath, $shapePlaceholder);

        $html .= '</'.$elementTag.'>';

        return $html;
    }

    /**
     * Transform p:graphicFrame tag
     *
     * @param DOMNode $childNode
     * @param DOMDocument $slideDom
     * @param DOMDocument $slideRelsDom
     * @param DOMDocument $slideLayoutDom
     * @param DOMDocument $slideLayoutRelsDom
     * @param DOMDocument $slideMasterDom
     * @return string
     */
    protected function transformP_GRAPHICFRAME($childNode, $slideDom, $slideRelsDom, $slideLayoutDom = null, $slideLayoutRelsDom = null, $slideMasterDom = null)
    {
        $nodeClass = $this->htmlPlugin->generateClassName();
        $this->css['.' .$nodeClass] = '';
        $html = '';

        $slideXPath = new DOMXPath($slideDom);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

        $slideRelsXPath = new DOMXPath($slideRelsDom);
        $slideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        $slideTables = $slideXPath->query('.//a:tbl', $childNode);

        if ($slideTables->length > 0) {
            $elementTag = $this->htmlPlugin->getTag('table');
        } else {
            // no table
            return '';
        }

        // get shape information
        list($shapeDrawing, $shapePlaceholder, $slideLayoutXPath, $slideMasterXPath) = $this->getShapeDetails($childNode, $slideXPath, $slideLayoutDom, $slideMasterDom);
        $this->css['.' . $nodeClass] .= 'position: absolute;top:' . $shapeDrawing['y'] . ';left:' . $shapeDrawing['x'] . ';width:' . $shapeDrawing['width'] . ';height:' . $shapeDrawing['height'] . ';';
        $this->css['.' . $nodeClass] .= 'table-layout:fixed;';
        if (isset($shapeDrawing['rotation'])) {
            $this->css['.' . $nodeClass] .= 'transform: rotate(' . $shapeDrawing['rotation'] . ');';
        }

        // handle styles

        // default styles
        $this->css['.' .$nodeClass] .= 'border: 1px solid black;border-collapse: collapse;';
        $this->css['.' .$nodeClass . ' th, td'] = '';
        $this->css['.' .$nodeClass . ' th, td'] .= 'border: 1px solid black;';

        // tblPr styles
        $nodesTblPr = $slideTables->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'tblPr');
        $nodeClassTableStyle = '';
        if ($nodesTblPr->length > 0) {
            // default table style name
            $nodesTableStyleId = $nodesTblPr->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'tableStyleId');
            if ($nodesTableStyleId->length > 0 && !empty($nodesTableStyleId->item(0)->nodeValue)) {
                $nodeClassTableStyle = ' t' . str_replace(array('{', '}'), '', $nodesTableStyleId->item(0)->nodeValue);
            }
        }

        // cell widths
        $nodesGridcol = $slideTables->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'gridCol');
        $tcWidths = array();
        foreach ($nodesGridcol as $nodeGridCol) {
            if ($nodeGridCol->hasAttribute('w')) {
                $tcWidths[] = $this->htmlPlugin->transformSizes($nodeGridCol->getAttribute('w'), 'twips');
            }
        }

        // handle contents
        $html .= '<' . $elementTag . ' class="' . $nodeClass . ' ' . (!empty($nodeClassTableStyle) ? $nodeClassTableStyle : '') . $this->htmlPlugin->getExtraClass('table') . '">';

        $nodesTr = $slideTables->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'tr');
        foreach ($nodesTr as $nodeTr) {
            // add rows
            $nodeClassTr = $this->htmlPlugin->generateClassName();
            $this->css['table.' . $nodeClass . ' tr.' .$nodeClassTr] = '';

            // handle styles
            if ($nodeTr->hasAttribute('h')) {
                $this->css['table.' . $nodeClass . ' tr.' .$nodeClassTr] .= 'height:' . $this->htmlPlugin->transformSizes($nodeTr->getAttribute('h'), 'twips') . ';';
            }

            // handle contents
            $html .= '<' . $this->htmlPlugin->getTag('tr') . ' class="' . $nodeClassTr . ' ' . $this->htmlPlugin->getExtraClass('tr') . '">';
            $nodesTc = $nodeTr->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'tc');
            $iNodeTc = 0; // keep track of the tc index to apply the width
            foreach ($nodesTc as $nodeTc) {
                // add cells
                $nodeClassTc = $this->htmlPlugin->generateClassName();
                $this->css['table.' . $nodeClass . ' tr td.' .$nodeClassTc] = '';

                // handle styles
                // width
                if (isset($tcWidths[$iNodeTc])) {
                    $this->css['table.' . $nodeClass . ' tr td.' .$nodeClassTc] .= 'width:' . $tcWidths[$iNodeTc] . ';';
                }

                // background color
                $nodesSolidFill = $nodeTc->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'solidFill');
                foreach ($nodesSolidFill as $nodeSolidFill) {
                    if ($nodeSolidFill->parentNode->tagName == 'a:tcPr') {
                        $this->css['table.' . $nodeClass . ' tr td.' .$nodeClassTc] .= 'background-color: #' . $this->getColor($nodeSolidFill) . ';';
                    }
                }

                // handle contents
                $html .= '<' . $this->htmlPlugin->getTag('tc') . ' class="' . $nodeClassTc . ' ' . $this->htmlPlugin->getExtraClass('tc') . '">';
                $html .= $this->transformP_TXBODY($nodeTc, $nodeClass, $slideDom, $slideXPath, $slideRelsDom, $slideRelsXPath, $slideLayoutDom, $slideLayoutXPath, $slideMasterDom, $slideMasterXPath, $shapePlaceholder);
                $html .= '</' . $this->htmlPlugin->getTag('tc') . '>';

                $iNodeTc++;
            }
            $html .= '</' . $this->htmlPlugin->getTag('tr') . '>';
        }

        $html .= '</'.$elementTag.'>';

        return $html;
    }

    /**
     * Transform p:pic tag
     *
     * @param DOMNode $childNode
     * @param DOMDocument $slideDom
     * @param DOMDocument $slideRelsDom
     * @param DOMDocument $slideLayoutDom
     * @param DOMDocument $slideLayoutRelsDom
     * @param DOMDocument $slideMasterDom
     * @return string
     */
    protected function transformP_PIC($childNode, $slideDom, $slideRelsDom, $slideLayoutDom = null, $slideLayoutRelsDom = null, $slideMasterDom = null)
    {
        $nodeClass = $this->htmlPlugin->generateClassName();
        $this->css['.' . $nodeClass] = '';
        $html = '';

        $slideXPath = new DOMXPath($slideDom);
        $slideXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $slideXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

        $slideRelsXPath = new DOMXPath($slideRelsDom);
        $slideRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        $slideImages = $slideXPath->query('.//a:blip[@r:embed]', $childNode);
        $slideAudios = $slideXPath->query('.//a:audioFile[@r:link]', $childNode);
        $slideVideos = $slideXPath->query('.//a:videoFile[@r:link]', $childNode);

        if ($slideAudios->length > 0) {
            $elementTag = $this->htmlPlugin->getTag('audio');
            $contentTag = 'audio';
            $nodeId = $slideAudios->item(0)->getAttribute('r:link');
        } elseif ($slideVideos->length > 0) {
            $elementTag = $this->htmlPlugin->getTag('video');
            $contentTag = 'video';
            $nodeId = $slideVideos->item(0)->getAttribute('r:link');
        } elseif ($slideImages->length > 0) {
            // audio and video elements may have cover images, ignore images in this case
            $elementTag = $this->htmlPlugin->getTag('image');
            $contentTag = 'image';
            $nodeId = $slideImages->item(0)->getAttribute('r:embed');
        }

        if (!isset($nodeId) || empty($nodeId) || !isset($contentTag) || !isset($elementTag) || empty($elementTag)) {
            return '';
        }

        // get shape information
        list($shapeDrawing, $shapePlaceholder, $slideLayoutXPath, $slideMasterXPath) = $this->getShapeDetails($childNode, $slideXPath, $slideLayoutDom, $slideMasterDom);
        $this->css['.' . $nodeClass] .= 'position: absolute;top:' . $shapeDrawing['y'] . ';left:' . $shapeDrawing['x'] . ';width:' . $shapeDrawing['width'] . ';height:' . $shapeDrawing['height'] . ';';
        if (isset($shapeDrawing['rotation'])) {
            $this->css['.' . $nodeClass] .= 'transform: rotate(' . $shapeDrawing['rotation'] . ');';
        }
        if (isset($shapeDrawing['border'])) {
            $this->css['.' . $nodeClass] .= 'border: ' . $shapeDrawing['border']['size'] . ' solid ' . $shapeDrawing['border']['color'] . ';';
        }

        $nodeRelationship = $slideRelsXPath->query('.//xmlns:Relationship[@Id="'.$nodeId.'"]');
        $target = str_replace('../', 'ppt/', $nodeRelationship->item(0)->getAttribute('Target'));
        $fileInfo = pathinfo($target);
        $ext = $fileInfo['extension'];
        $fileString = $this->pptxStructure->getContent($target);

        if ($this->htmlPlugin->getFilesAsBase64()) {
            $src = 'data:' . $contentTag . '/' . $ext . ';base64,' . base64_encode($fileString);
        } else {
            $src = $this->htmlPlugin->getOutputFilesPath() . $fileInfo['filename'] . '.' . $ext;
            file_put_contents($src, $fileString);
        }

        // handle styles
        // get descr
        $extraAttributes = '';
        $nodesCNvPr = $slideXPath->query('.//p:nvPicPr/p:cNvPr', $childNode);
        if ($nodesCNvPr->length > 0 && $nodesCNvPr->item(0)->hasAttribute('descr')) {
            $extraAttributes = 'descr="' . $nodesCNvPr->item(0)->getAttribute('descr') . '"';
        }

        // handle contents
        if ($contentTag == 'image') {
            $isLink = false;

            // check if link
            $nodesHlinkClick = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'hlinkClick');
            if ($nodesHlinkClick->length > 0 && $nodesHlinkClick->item(0)->hasAttribute('r:id')) {
                $isLink = true;
                // get target link from rels
                $nodesTarget = $slideRelsXPath->query('//xmlns:Relationship[@Id="' . $nodesHlinkClick->item(0)->getAttribute('r:id') . '"]');
                if ($nodesTarget->length > 0) {
                    $html .= '<' . $this->htmlPlugin->getTag('hyperlink') . ' href="' . $nodesTarget->item(0)->getAttribute('Target') . '" target="_blank">';
                }
            }

            $html .= '<'.$elementTag.' class="'.$nodeClass.' ' . $this->htmlPlugin->getExtraClass($contentTag) . '" src="' . $src . '" ' . $extraAttributes . '>';
            $html .= '</'.$elementTag.'>';

            // close link tag
            if ($isLink) {
                $html .= '</' . $this->htmlPlugin->getTag('hyperlink') . '>';
            }
        } elseif ($contentTag == 'audio' || $contentTag == 'video') {
            $html .= '<'.$elementTag.' class="'.$nodeClass.' ' . $this->htmlPlugin->getExtraClass($contentTag) . '" src="' . $src . '" controls ' . $extraAttributes . '>';
            $html .= '</'.$elementTag.'>';
        }

        return $html;
    }

    /**
     * Transform p:txBody tag
     *
     * @param DOMNode $childNode
     * @param string $nodeClassParent
     * @param DOMDocument $slideDom
     * @param DOMXPath $slideXPath
     * @param DOMDocument $slideRelsDom
     * @param DOMXPath $slideRelsXPath
     * @param DOMDocument $slideLayoutDom
     * @param DOMXPath $slideLayoutXPath
     * @param DOMDocument $slideMasterDom
     * @param DOMXPath $slideMasterXPath
     * @param array $shapePlaceholder
     * @return string
     */
    protected function transformP_TXBODY($childNode, $nodeClassParent, $slideDom, $slideXPath, $slideRelsDom, $slideRelsXPath, $slideLayoutDom, $slideLayoutXPath, $slideMasterDom, $slideMasterXPath, $shapePlaceholder)
    {
        $html = '';

        // default element
        $elementTagParagraph = $this->htmlPlugin->getTag('paragraph');
        $elementTagSpan = $this->htmlPlugin->getTag('span');

        $this->css['.' . $nodeClassParent . ' ' . $elementTagParagraph] = '';
        $this->css['.' . $nodeClassParent . ' ' . $elementTagSpan] = '';

        // handle txBody styles
        // get styles them from the slide master
        if (!is_null($slideMasterDom) && !is_null($slideMasterXPath) && isset($shapePlaceholder['type']) && !is_null($shapePlaceholder['type'])) {
            $this->setCssStylesMaster($slideMasterDom, $slideMasterXPath, $shapePlaceholder['type'], $nodeClassParent, $elementTagParagraph);
        }
        if (!is_null($slideMasterDom) && !is_null($slideMasterXPath) && count($shapePlaceholder) > 0) {
            $this->setCssStylesTxBody($slideMasterDom, $slideMasterXPath, $shapePlaceholder, $nodeClassParent, $elementTagParagraph, array('idx', 'type'));
        }

        // get styles them from the slide layout
        if (!is_null($slideLayoutDom) && !is_null($slideLayoutXPath) && count($shapePlaceholder) > 0) {
            $this->setCssStylesTxBody($slideLayoutDom, $slideLayoutXPath, $shapePlaceholder, $nodeClassParent, $elementTagParagraph, array('idx', 'type'));
        }

        // get styles from the shape
        if (isset($shapePlaceholder['id']) && !is_null($shapePlaceholder['id'])) {
            $this->setCssStylesTxBody($slideDom, $slideXPath, array('id' => $shapePlaceholder['id']), $nodeClassParent, $elementTagParagraph, array('id'));
        }

        // handle paragraphs
        $nodesP = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'p');
        foreach ($nodesP as $nodeP) {
            $nodeClassParagraph = $this->htmlPlugin->generateClassName();
            $this->css[$elementTagParagraph . '.' . $nodeClassParagraph] = '';

            // paragraph styles
            $nodesPpr = $nodeP->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'pPr');
            if ($nodesPpr->length > 0) {
                $this->css[$elementTagParagraph . '.' . $nodeClassParagraph] .= $this->getPprStyles($nodesPpr->item(0));
            }

            $html .= '<'.$elementTagParagraph.' class="'.$nodeClassParagraph.' ' . $this->htmlPlugin->getExtraClass('paragraph') . '">';

            if ($nodeP->hasChildNodes()) {
                foreach ($nodeP->childNodes as $childNode) {
                    switch ($childNode->nodeName) {
                        case 'a:br':
                            $html .= '<' . $this->htmlPlugin->getTag('break') . '>';
                            break;
                        case 'a:r':
                            $nodeClassSpan = $this->htmlPlugin->generateClassName();
                            $this->css[$elementTagSpan . '.' . $nodeClassSpan] = '';

                            // span styles
                            $isLink = false;
                            $nodesRpr = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'rPr');
                            if ($nodesRpr->length > 0) {
                                $this->css[$elementTagSpan . '.' . $nodeClassSpan] .= $this->getRprStyles($nodesRpr->item(0));

                                // check if link
                                $nodesHlinkClick = $nodesRpr->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'hlinkClick');
                                if ($nodesHlinkClick->length > 0 && $nodesHlinkClick->item(0)->hasAttribute('r:id') && !empty($nodesHlinkClick->item(0)->getAttribute('r:id'))) {
                                    $isLink = true;
                                    // get target link from rels
                                    $nodesTarget = $slideRelsXPath->query('//xmlns:Relationship[@Id="' . $nodesHlinkClick->item(0)->getAttribute('r:id') . '"]');
                                    if ($nodesTarget->length > 0) {
                                        if ($nodesHlinkClick->item(0)->hasAttribute('action') && strpos($nodesHlinkClick->item(0)->getAttribute('action'), 'ppaction://hlinksldjump') !== false) {
                                            // internal link
                                            $html .= '<' . $this->htmlPlugin->getTag('hyperlink') . ' href="#' . str_replace('.', '_', $nodesTarget->item(0)->getAttribute('Target')) . '">';
                                        } else {
                                            // external link
                                            $html .= '<' . $this->htmlPlugin->getTag('hyperlink') . ' href="' . $nodesTarget->item(0)->getAttribute('Target') . '" target="_blank">';
                                        }
                                    }
                                }
                            }

                            $html .= '<'.$elementTagSpan.' class="'.$nodeClassSpan.' ' . $this->htmlPlugin->getExtraClass('span') . '">';

                            // handle text
                            $nodesT = $childNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't');
                            foreach ($nodesT as $nodeT) {
                                $html .= $this->transformW_T($nodeT);
                            }

                            // close span tag
                            $html .= '</'.$elementTagSpan.'>';

                            // close link tag
                            if ($isLink) {
                                $html .= '</' . $this->htmlPlugin->getTag('hyperlink') . '>';
                            }
                            break;
                        default:
                            break;
                    }
                }
            }

            $html .= '</'.$elementTagParagraph.'>';
        }

        return $html;
    }

    /**
     * Transform w:t tag
     *
     * @param DOMNode $childNode
     * @return string
     */
    protected function transformW_T($childNode)
    {
        // fix < and > values
        $html = str_replace(array('<', '>'), array('&lt;', '&gt;'), $childNode->nodeValue);

        return $html;
    }
}
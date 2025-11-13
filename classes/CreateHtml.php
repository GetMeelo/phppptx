<?php

/**
 * Create html
 *
 * @category   Phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
require_once __DIR__ . '/DOMPDF_lib.php';

 class CreateHtml extends CreateElement
{
    /**
     *
     * @access protected
     * @var array
     */
    protected $blockContents;

    /**
     *
     * @access protected
     * @var array
     */
    protected $blockStyles;

    /**
     *
     * @access public
     * @static
     * @var bool
     */
    public static $htmlExtended;

    /**
     *
     * @access protected
     * @var string
     */
    protected $linkElement;

    /**
     *
     * @access protected
     * @var array
     */
    protected $listElement;

    /**
     *
     * @access public
     * @static
     * @var bool
     */
    public static $parseCSSVars;

    /**
     *
     * @access protected
     * @var array
     */
    protected $positionHtml;

    /**
     *
     * @access public
     * @static
     * @var CreatePptx
     */
    public $pptx;

    /**
     *
     * @access protected
     * @var string
     */
    protected $xml;

    /**
     * Constructor
     *
     * @param CreatePptx $pptx
     */
    public function __construct($pptx)
    {
        $this->blockContents = array();
        $this->blockStyles = array();
        $this->linkElement = '';
        $this->listElement = array();
        self::$parseCSSVars = false;
        self::$htmlExtended = false;
        $this->pptx = $pptx;
    }

    /**
     * Creates and adds HTML
     *
     * @access public
     * @param string $html
     *      HTML standard
     *          <a>
     *          <p>, <h1>, <h2>, <h3>, <h4>, <h5>, <h6>         : background-color, rtl, text-align (left, center, right, justify)
     *          <span>, #text                                   : background-color, color, font-family, font-size, font-style (italic, oblique), font-weight (bold, bolder, 700, 800, 900), letter-spacing, text-decoration (line-through, underline)
     *          <b>, <cite>, <em>, <i>, <s>, <strong>, <var>
     *          <br>
     *          <ul>, <ol>                                      : list-style-type (disc, circle, square, decimal, lower-alpha, lower-latin, lower-roman, upper-alpha, upper-latin, upper-roman, and PowerPoint list style names)
     *          <li>
     *      HTML Extended
     *          phppptx_audio
     *          phppptx_image
     *          phppptx_slide
     *          phppptx_svg
     *          phppptx_video
     * @param array $options
     *      'disableWrapValue' (bool) if true disable using a wrap value with Tidy. Default as false
     *      'forceNotTidy' (bool) if true, avoid using Tidy. Only recommended if Tidy can't be installed. Default as false
     *      'insertMode' (string) replace, ignore. Default as replace
     *      'parseCSSVars' (bool) parse CSS variables. Default as false
     *      'positionHtml' (string) position to be used in HTML Extended tags
     *      'useHtmlExtended' (bool) use HTML Extended and CSS Extended. Default as false
     */
    public function createElementHtml($html, $options)
    {
        if (isset($options['parseCSSVars'])) {
            self::$parseCSSVars = $options['parseCSSVars'];
        }
        if (isset($options['positionHtml'])) {
            $this->positionHtml = $options['positionHtml'];
        }
        if (isset($options['useHtmlExtended'])) {
            self::$htmlExtended = $options['useHtmlExtended'];
        }

        $dompdf = new PARSERHTML();
        $dompdfTree = $dompdf->getDompdfTree($html, $options);

        $this->renderHtml($dompdfTree);

        // add remaining content if some exists
        $this->getBlockContent('paragraph');

        return $this->xml;
    }

    /**
     * Renders HTML
     *
     * @param array $node
     */
    protected function renderHtml($node, $depth = 0)
    {
        $nodeAttributes = isset($node['attributes']) ? $node['attributes'] : array();
        $nodeProperties = isset($node['properties']) ? $node['properties'] : array();

        switch ($node['nodeName']) {
            case 'h1':
            case 'h2':
            case 'h3':
            case 'h4':
            case 'h5':
            case 'h6':
            case 'li':
            case 'p':
                // block content

                // generate current block content
                $this->getBlockContent('paragraph');

                // add the styles
                $this->blockStyles = array();

                if (isset($nodeProperties['text_align'])) {
                    // textAlign
                    $horizontalAlign = $this->normalizeHorizontalAlign($nodeProperties['text_align']);

                    $this->blockStyles['align'] = $horizontalAlign;
                }

                if ((isset($nodeAttributes['dir']) && strtolower($nodeAttributes['dir']) == 'rtl') || (isset($nodeProperties['direction']) && strtolower($nodeProperties['direction']) == 'rtl') || CreatePptx::$rtl) {
                    // rtl
                    $this->blockStyles['rtl'] = true;
                }

                if ($node['nodeName'] == 'li' && isset($this->listElement[$depth - 1])) {
                    // list
                    $this->blockStyles['listLevel'] = array_search($depth - 1, array_keys($this->listElement));
                    $this->blockStyles['listStyles'] = $this->listElement[$depth - 1];
                }

                break;
            case 'ul':
            case 'ol':
                $listType = 'filledRoundBullet';
                if ($node['nodeName'] == 'ol') {
                    $listType = 'arabicPeriod';
                }

                if (isset($nodeProperties['list_style_type'])) {
                    switch ($nodeProperties['list_style_type']) {
                        case 'disc':
                            $listType = 'filledRoundBullet';
                            break;
                        case 'circle':
                            $listType = 'hollowRoundBullet';
                            break;
                        case 'square':
                            $listType = 'filledSquareBullet';
                            break;
                        case 'decimal':
                            $listType = 'arabicPeriod';
                            break;
                        case 'lower-alpha':
                        case 'lower-latin':
                            $listType = 'alphaLcPeriod';
                            break;
                        case 'lower-roman':
                            $listType = 'romanLcPeriod';
                            break;
                        case 'upper-alpha':
                        case 'upper-latin':
                            $listType = 'alphaUcPeriod';
                            break;
                        case 'upper-roman':
                            $listType = 'romanUcPeriod';
                            break;
                        case 'filledroundbullet':
                            $listType = 'filledRoundBullet';
                            break;
                        case 'hollowroundbullet':
                            $listType = 'hollowRoundBullet';
                            break;
                        case 'filledsquarebullet':
                            $listType = 'filledSquareBullet';
                            break;
                        case 'hollowsquarebullet':
                            $listType = 'hollowSquareBullet';
                            break;
                        case 'starbullet':
                            $listType = 'starBullet';
                            break;
                        case 'arrowbullet':
                            $listType = 'arrowBullet';
                            break;
                        case 'checkmarkbullet':
                            $listType = 'checkmarkBullet';
                            break;
                        case 'romanuppercase':
                            $listType = 'romanUpperCase';
                            break;
                        case 'romanlowercase':
                            $listType = 'romanLowerCase';
                            break;
                        case 'alphauppercase':
                            $listType = 'alphaUpperCase';
                            break;
                        case 'alphalowercase':
                            $listType = 'alphaLowerCase';
                            break;
                        case 'alphalcparenboth':
                            $listType = 'alphaLcParenBoth';
                            break;
                        case 'alphaucparenboth':
                            $listType = 'alphaUcParenBoth';
                            break;
                        case 'alphalcparenr':
                            $listType = 'alphaLcParenR';
                            break;
                        case 'alphaucparenr':
                            $listType = 'alphaUcParenR';
                            break;
                        case 'alphalcperiod':
                            $listType = 'alphaLcPeriod';
                            break;
                        case 'alphaucperiod':
                            $listType = 'alphaUcPeriod';
                            break;
                        case 'arabicparenboth':
                            $listType = 'arabicParenBoth';
                            break;
                        case 'arabicparenr':
                            $listType = 'arabicParenR';
                            break;
                        case 'arabicperiod':
                            $listType = 'arabicPeriod';
                            break;
                        case 'arabicplain':
                            $listType = 'arabicPlain';
                            break;
                        case 'romanlcparenboth':
                            $listType = 'romanLcParenBoth';
                            break;
                        case 'romanuparenboth':
                            $listType = 'romanUcParenBoth';
                            break;
                        case 'romanlcparenr':
                            $listType = 'romanLcParenR';
                            break;
                        case 'romanuparenr':
                            $listType = 'romanUcParenR';
                            break;
                        case 'romanlcperiod':
                            $listType = 'romanLcPeriod';
                            break;
                        case 'romanuperiod':
                            $listType = 'romanUcPeriod';
                            break;
                        case 'circlenumdbplain':
                            $listType = 'circleNumDbPlain';
                            break;
                        case 'circlenumwdblackplain':
                            $listType = 'circleNumWdBlackPlain';
                            break;
                        case 'circlenumwdwhiteplain':
                            $listType = 'circleNumWdWhitePlain';
                            break;
                        case 'arabicdbperiod':
                            $listType = 'arabicDbPeriod';
                            break;
                        case 'arabicdbplain':
                            $listType = 'arabicDbPlain';
                            break;
                        case 'ea1chsperiod':
                            $listType = 'ea1ChsPeriod';
                            break;
                        case 'ea1chsplain':
                            $listType = 'ea1ChsPlain';
                            break;
                        case 'ea1chtperiod':
                            $listType = 'ea1ChtPeriod';
                            break;
                        case 'ea1chtplain':
                            $listType = 'ea1ChtPlain';
                            break;
                        case 'ea1jpnchsdbperiod':
                            $listType = 'ea1JpnChsDbPeriod';
                            break;
                        case 'ea1jpnkorplain':
                            $listType = 'ea1JpnKorPlain';
                            break;
                        case 'ea1jpnkorperiod':
                            $listType = 'ea1JpnKorPeriod';
                            break;
                        case 'arabic1minus':
                            $listType = 'arabic1Minus';
                            break;
                        case 'arabic2minus':
                            $listType = 'arabic2Minus';
                            break;
                        case 'hebrew2minus':
                            $listType = 'hebrew2Minus';
                            break;
                        case 'thaialphaperiod':
                            $listType = 'thaiAlphaPeriod';
                            break;
                        case 'thaialphaparenr':
                            $listType = 'thaiAlphaParenR';
                            break;
                        case 'thaialphaparenboth':
                            $listType = 'thaiAlphaParenBoth';
                            break;
                        case 'thainumperiod':
                            $listType = 'thaiNumPeriod';
                            break;
                        case 'thainumparenr':
                            $listType = 'thaiNumParenR';
                            break;
                        case 'thainumparenboth':
                            $listType = 'thaiNumParenBoth';
                            break;
                        case 'hindialphaperiod':
                            $listType = 'hindiAlphaPeriod';
                            break;
                        case 'hindinumperiod':
                            $listType = 'hindiNumPeriod';
                            break;
                        case 'hindinumparenr':
                            $listType = 'hindiNumParenR';
                            break;
                        case 'hindialpha1period':
                            $listType = 'hindiAlpha1Period';
                            break;
                        default:
                            $listType = 'filledRoundBullet';
                            break;
                    }
                }

                $this->listElement[$depth] = array(
                    'type' => $listType,
                );

                break;
            case 'a':
                // hyperlink

                if (isset($node['attributes']['href']) && $node['attributes']['href'] != '') {
                    $this->linkElement = $node['attributes']['href'];
                }

                break;
            case 'br':
                // break
                $this->blockContents[] = "\n";
                $this->blockStyles['parseLineBreaks'] = true;

                break;
            case '#text':
                // add the text contents
                $inlineContents = array();
                if (isset($nodeProperties['background_color']) && is_array($nodeProperties['background_color'])) {
                    // backgroundColor
                    $color = $this->normalizeColor($nodeProperties['background_color']);

                    $inlineContents['highlight'] = $color;
                }
                if (isset($nodeProperties['color']) && !empty($nodeProperties['color']) && is_array($nodeProperties['color'])) {
                    // color
                    $color = $this->normalizeColor($nodeProperties['color']);

                    $inlineContents['color'] = $color;
                }
                if (isset($nodeProperties['font_family']) && $nodeProperties['font_family'] != 'serif') {
                    // font
                    $arrayFonts = explode(',', $nodeProperties['font_family']);
                    $font = trim($arrayFonts[0]);
                    $font = str_replace(array('"', "'"), '', $font);

                    $inlineContents['font'] = $font;
                }
                if (isset($nodeProperties['font_size']) && !empty($nodeProperties['font_size'])) {
                    // fontSize
                    $inlineContents['fontSize'] = round($nodeProperties['font_size']);
                }
                if (isset($nodeProperties['font_style']) && ($nodeProperties['font_style'] == 'italic' || $nodeProperties['font_style'] == 'oblique')) {
                    // italic
                    $inlineContents['italic'] = true;
                }
                if (isset($nodeProperties['font_weight']) && ($nodeProperties['font_weight'] == 'bold' || $nodeProperties['font_weight'] == 'bolder' || $nodeProperties['font_weight'] == '700' || $nodeProperties['font_weight'] == '800' || $nodeProperties['font_weight'] == '900')) {
                    // bold
                    $inlineContents['bold'] = true;
                }
                if (isset($nodeProperties['letter_spacing']) && !empty($nodeProperties['letter_spacing'])) {
                    // letterSpacing
                    $inlineContents['characterSpacing'] = round($nodeProperties['letter_spacing']);
                }
                if (isset($nodeProperties['text_decoration']) && $nodeProperties['text_decoration'] == 'line-through') {
                    // strikethrough
                    $inlineContents['strikethrough'] = 'single';
                }
                if (isset($nodeProperties['text_decoration']) && $nodeProperties['text_decoration'] == 'underline') {
                    // underline
                    $inlineContents['underline'] = 'single';
                }
                if (isset($node['nodeValue'])) {
                    // text content
                    // clean extra line feeds generated by the parser after <br /> tags that may give problems with the rendering
                    $inlineContents['text'] = preg_replace('/[\n\r]+/', '', $node['nodeValue']);
                }
                if (!empty($this->linkElement)) {
                    // link
                    $inlineContents['hyperlink'] = $this->linkElement;
                    $this->linkElement = '';
                }

                $this->blockContents[] = $inlineContents;

                break;
            default:
                if (file_exists(__DIR__ . '/HTMLExtended.php') && self::$htmlExtended && $this->pptx) {
                    // support extended tags
                    $htmlExtended = new HTMLExtended();
                    $extendedTags = HTMLExtended::getTagsInline() + HTMLExtended::getTagsBlock();

                    if (array_key_exists($node['nodeName'], $extendedTags)) {
                        $attributesContent = array();
                        foreach ($node['attributes'] as $key => $value ) {
                            $attributesContent[str_replace('data-', '', $key)] = $value;
                        }

                        // normalize attributes names
                        $attributesContent = $this->normalizeAttributesNames($attributesContent);

                        // normalize attributes values
                        $attributesContent = $this->normalizeAttributesValues($attributesContent);

                        if (!isset($attributesContent['position'])) {
                            // set the position if not set
                            $attributesContent['position'] = $this->positionHtml;
                        }

                        if (isset($attributesContent['position']['new'])) {
                            $attributesContent['position'] = $attributesContent['position']['new'];
                        }

                        if ($extendedTags[$node['nodeName']] == 'addAudio') {
                            $this->pptx->addAudio($attributesContent['src'], $attributesContent['position'], $attributesContent, $attributesContent);
                        } elseif ($extendedTags[$node['nodeName']] == 'addImage') {
                            $this->pptx->addImage($attributesContent['src'], $attributesContent['position'], $attributesContent, $attributesContent);
                        } elseif ($extendedTags[$node['nodeName']] == 'addSlide') {
                            $this->pptx->addSlide($attributesContent);
                        } elseif ($extendedTags[$node['nodeName']] == 'addSvg') {
                            $this->pptx->addSvg($attributesContent['src'], $attributesContent['position'], $attributesContent, $attributesContent);
                        } elseif ($extendedTags[$node['nodeName']] == 'addVideo') {
                            $this->pptx->addVideo($attributesContent['src'], $attributesContent['position'], $attributesContent, $attributesContent);
                        }
                    }
                }
                break;
        }
        if (isset($nodeProperties['display']) && $nodeProperties['display'] == 'none') {
            //do not render the subtree
        } else {
            if (!empty($node['children'])) {
                $depth++;
                foreach ($node['children'] as $child) {
                    $this->renderHtml($child, $depth);
                }
            }
        }
    }

    /**
     * Generates block content
     *
     * @access private
     * @param string $type paragraph
     */
    private function getBlockContent($type)
    {
        if ($type == 'paragraph') {
            $text = new CreateText();
            $newContentText = $text->createElementText($this->blockContents, $this->blockStyles);

            // handle external relationships such as hyperlinks
            $externalRelationships = $text->getExternalRelationships();
            $this->externalRelationships = array_merge($this->externalRelationships, $externalRelationships);

            $this->xml .= $newContentText;

            $this->blockContents = array();
            $this->blockStyles = array();
        }
    }

    /**
     * Normalizes a color
     *
     * @access private
     * @param array $color
     * @return string
     */
    private function normalizeColor($color)
    {
        if (strtolower($color['hex']) == 'transparent') {
            return '';
        } else {
            return strtoupper(str_replace('#', '', $color['hex']));
        }
    }

    /**
     * Normalizes a horizontal align value
     *
     * @access private
     * @param string $horizontalAlign Horizontal align
     * @return string
     */
    private function normalizeHorizontalAlign($horizontalAlign = 'left')
    {
        $normalizedHorizontalAlign = $horizontalAlign;

        switch ($horizontalAlign) {
            case 'left':
            case 'justify':
                $normalizedHorizontalAlign = 'left';
                break;
            case 'center':
                $normalizedHorizontalAlign = 'center';
                break;
            case 'right':
                $normalizedHorizontalAlign = 'right';
                break;
            default:
                $normalizedHorizontalAlign = 'left';
        }

        return $normalizedHorizontalAlign;
    }

    /**
     * Normalize attribute names
     *
     * @param array $attributes
     * @return array
     */
    private function normalizeAttributesNames($attributes)
    {
        $attributesNormalized = array();

        // get the correct attribute setting upper and lower cases
        foreach ($attributes as $key => $value) {
            $keyInitial = $key;

            if ($keyInitial === 'cleanlayoutparagraphcontents') {
                $keyInitial = 'cleanLayoutParagraphContents';
            }

            if ($keyInitial === 'cleanslideplaceholdertypes') {
                $keyInitial = 'cleanSlidePlaceholderTypes';
            }

            $attributesNormalized[$keyInitial] = $value;
        }

        return $attributesNormalized;
    }

    /**
     * Normalize attribute values
     *
     * @param array $attributes
     * @return array
     */
    private function normalizeAttributesValues($attributes)
    {
        // replace true and false strings by boolean values and JSON
        $attributesNormalized = array();

        $booleanTags = array();
        $mixedTags = array();
        $jsonTags = array('cleanSlidePlaceholderTypes', 'position');

        foreach ($attributes as $key => $value) {
            $valueInitial = $value;

            if (in_array($key, $booleanTags)) {
                $valueInitial = ($value == 'true') ? true : false;
            }

            if (in_array($key, $jsonTags)) {
                $valueInitial = json_decode($value, true);
            }

            if (in_array($key, $mixedTags)) {
                if ($value == 'true') {
                    $valueInitial = true;
                } else if ($value == 'false') {
                    $valueInitial = false;
                } else {
                    $valueInitial = $value;
                }
            }

            $attributesNormalized[$key] = $valueInitial;
        }

        return $attributesNormalized;
    }
}
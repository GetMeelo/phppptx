<?php

/**
 * Transform PPTX to HTML using native PHP classes. Default plugin
 *
 * @category   Phppptx
 * @package    transform
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class TransformNativeHtmlDefaultPlugin extends TransformNativeHtmlPlugin
{
    /**
     * Conversion factor, used by the transformSizes method with px as target
     *
     * @var float
     */
    protected $conversionFactor = 1.3;

    /**
     * Add file src as base64
     *
     * @var boolean
     */
    protected $filesAsBase64 = true;

    /**
     * Target folder for files and other external contents. Not used for files is $filesAsBase64 is true
     *
     * @var string
     */
    protected $outputFilesPath = 'output_files/';

    /**
     * Scale factor, used by the transformSizes method
     *
     * @var float
     */
    protected $scaleFactor = 600;

    /**
     * Default conversion unit: pt, px
     *
     * @var string
     */
    protected $unit = 'pt';

    /**
     * OOXML => HTML tags
     *
     * @var array
     */
    protected $tags = array(
        'audio' => 'audio',
        'break' => 'br',
        'hyperlink' => 'a',
        'image' => 'img',
        'paragraph' => 'p',
        'shape' => 'div',
        'span' => 'span',
        'table' => 'table',
        'tc' => 'td',
        'tr' => 'tr',
        'video' => 'video',
    );

    /**
     * Constructor. Init HTML, CSS, meta and javascript base contents
     */
    public function __construct() {
        $this->baseCSS = '<style>.slide{border: 1pt solid; margin-top: 10pt;} span.tabcontent{margin-left: 40px;} p {margin-block-start:0; margin-block-end:0; margin-inline-start:0; margin-inline-end:0;padding-block-start:0; padding-block-end:0; padding-inline-start:0; padding-inline-end:0;}</style>';
        $this->baseHTML = '<!DOCTYPE html><html>';
        $this->baseJavaScript = '';
        $this->baseMeta = '<meta charset="UTF-8">';
    }

    /**
     * Generate class name to be added to tags
     *
     * @return string Class name
     */
    public function generateClassName()
    {
        return str_replace('.', '_', uniqid('pptx_', true));
    }

    /**
     * Transform colors
     *
     * @return string New color
     */
    public function transformColors($color) {
        $colorTarget = $color;

        if ($color == 'auto' || empty($color)) {
            $colorTarget = '000000';
        }

        return $colorTarget;
    }

    /**
     * Transform content sizes
     *
     * @param mixed $value OOXML size
     * @param string $source OOXML type size eights, fifths-percent, half-points, hundreds, init, pts, twips
     * @param string $target Target size deg, pt, px, %
     * @param bool $applyScaleFactor Apply scale factor
     * @return string HTML/CSS size
     */
    public function transformSizes($value, $source, $target = null, $applyScaleFactor = true)
    {
        $returnValue = 0;
        $value = (float)$value;

        if ($target === null) {
            $target = $this->unit;
        }

        if ($source == 'eights' && $value) {
            if ($target == 'pt') {
                $returnValue = ($value / 8);
            } elseif ($target == 'px') {
                $returnValue = ($value / 8) * $this->conversionFactor;
            }

            // minimum value
            if ($returnValue < 1) {
                $returnValue = 1;
            }
        }

        if ($source == 'fifths-percent' && $value) {
            if ($target == '%') {
                $returnValue = ($value / 50);
            }
        }

        if ($source == 'half-points' && $value) {
            if ($target == 'pt') {
                $returnValue = ($value / 2);
            } elseif ($target == 'px') {
                $returnValue = ($value / 2) * $this->conversionFactor;
            }
        }

        if ($source == 'hundreds' && $value) {
            if ($target == 'pt') {
                $returnValue = ($value / 100);
            } elseif ($target == 'px') {
                $returnValue = ($value / 100) * $this->conversionFactor;
            }
        }

        if ($source == 'init' && $value) {
            if ($target == 'pt') {
                $returnValue = $value;
            } elseif ($target == 'px') {
                $returnValue = $value * $this->conversionFactor;
            }

            // minimum value
            if ($returnValue < 1) {
                $returnValue = 1;
            }
        }

        if ($source == 'pts' && $value) {
            if ($target == 'px') {
                $returnValue = $value * $this->conversionFactor;
            }
        }

        if ($source == 'twips' && $value) {
            if ($target == 'deg') {
                $returnValue = ($value / 60000);
            } elseif ($target == 'pt') {
                $returnValue = ($value / 20);
            } elseif ($target == 'px') {
                $returnValue = ($value / 20) * $this->conversionFactor;
            }
        }

        // apply scale factor
        if ($this->scaleFactor > 0 && $applyScaleFactor) {
            $returnValue = $returnValue / $this->scaleFactor;
        }

        // normalize decimal values to use dots
        $returnValue = str_replace(',', '.', (string)$returnValue);

        return (string)$returnValue . $target;
    }
}

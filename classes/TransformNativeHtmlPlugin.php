<?php

/**
 * Transform PPTX to HTML using native PHP classes. Abstract class
 *
 * @category   Phppptx
 * @package    transform
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
abstract class TransformNativeHtmlPlugin
{
    /**
     * Base CSS
     * @var string
     */
    protected $baseCSS;

    /**
     * Base HTML
     * @var string
     */
    protected $baseHTML;

    /**
     * Base JavaScript
     * @var string
     */
    protected $baseJavaScript;

    /**
     * Base Meta
     * @var string
     */
    protected $baseMeta;

    /**
     * Conversion factor, used by the transformSizes method
     * @var float
     */
    protected $conversionFactor;

    /**
     * Extra classes
     * @var array
     */
    protected $extraClasses = array();

    /**
     * Files as base64
     * @var bool
     */
    protected $filesAsBase64;

    /**
     * Target folder for files and other external contents. Not used for files is $filesAsBase64 is true
     * @var string
     */
    protected $outputFilesPath;

    /**
     * Scale factor
     * @var float
     */
    protected $scaleFactor;

    /**
     * OOXML => HTML tags
     * @var array
     */
    protected $tags = array();

    /**
     * Conversion unit
     * @var mixed
     */
    protected $unit;

    /**
     * Generate class name to be added to tags
     */
    abstract public function generateClassName();

    /**
     * Transform colors
     * @param string $color
     * @return string New color
     */
    abstract public function transformColors($color);

    /**
     * Transform content sizes
     *
     * @param mixed $value OOXML size
     * @param string $source OOXML type size
     * @param string $target Target size
     * @param bool $applyScaleFactor Apply scale factor
     * @return string HTML/CSS size
     */
    abstract public function transformSizes($value, $source, $target = null, $applyScaleFactor = true);

    /**
     * Getter $baseCSS
     * @return string
     */
    public function getBaseCSS()
    {
        return $this->baseCSS;
    }
    /**
     * Setter $baseCSS
     * @param string $baseCSS
     */
    public function setBaseCSS($baseCSS)
    {
        $this->baseCSS = $baseCSS;
    }

    /**
     * Getter $baseHTML
     * @return string
     */
    public function getBaseHTML()
    {
        return $this->baseHTML;
    }
    /**
     * Setter $baseHTML
     * @param string $baseHTML
     */
    public function setBaseHTML($baseHTML)
    {
        $this->baseHTML = $baseHTML;
    }

    /**
     * Getter $baseJavaScript
     * @return string
     */
    public function getBaseJavaScript()
    {
        return $this->baseJavaScript;
    }
    /**
     * Setter $baseJavaScript
     * @param string $baseJavaScript
     */
    public function setBaseJavaScript($baseJavaScript)
    {
        $this->baseJavaScript = $baseJavaScript;
    }

    /**
     * Getter $baseMeta
     * @return string
     */
    public function getBaseMeta()
    {
        return $this->baseMeta;
    }
    /**
     * Setter $baseMeta
     * @param string $baseMeta
     */
    public function setBaseMeta($baseMeta)
    {
        $this->baseMeta = $baseMeta;
    }

    /**
     * Getter $conversionFactor
     * @return float
     */
    public function getConversionFactor()
    {
        return $this->conversionFactor;
    }
    /**
     * Setter $conversionFactor
     * @param float $conversionFactor
     */
    public function setConversionFactor($conversionFactor)
    {
        $this->conversionFactor = $conversionFactor;
    }

    /**
     * Getter extra class value
     * @return void|string
     */
    public function getExtraClass($tag)
    {
        if (isset($this->extraClasses[$tag])) {
            return $this->extraClasses[$tag];
        }
    }
    /**
     * Getter $extraClasses
     * @return array
     */
    public function getExtraClasses()
    {
        return $this->extraClasses;
    }
    /**
     * Setter $extraClasses
     * @param string $tag
     * @param string $class
     */
    public function setExtraClasses($tag, $class)
    {
        $this->extraClasses[$tag] = $class;
    }

    /**
     * Getter $filesAsBase64
     * @return bool
     */
    public function getFilesAsBase64()
    {
        return $this->filesAsBase64;
    }
    /**
     * Setter $filesAsBase64
     * @param bool $filesAsBase64
     */
    public function setFilesAsBase64($filesAsBase64)
    {
        $this->filesAsBase64 = $filesAsBase64;
    }

    /**
     * Getter $outputFilesPath
     * @return string
     */
    public function getOutputFilesPath()
    {
        return $this->outputFilesPath;
    }
    /**
     * Setter $outputFilesPath
     * @param string $outputFilesPath
     */
    public function setOutputFilesPath($outputFilesPath)
    {
        $this->outputFilesPath = $outputFilesPath;
    }

    /**
     * Getter $scaleFactor
     * @return float
     */
    public function getScaleFactor()
    {
        return $this->scaleFactor;
    }
    /**
     * Setter $scaleFactor
     * @param float $scaleFactor
     */
    public function setScaleFactor($scaleFactor)
    {
        $this->scaleFactor = $scaleFactor;
    }

    /**
     * Getter $tag value
     * @param string $tag
     * @return string
     */
    public function getTag($tag)
    {
        return $this->tags[$tag];
    }
    /**
     * Setter $setTag
     * @param string $tag
     * @param string $value
     */
    public function setTag($tag, $value)
    {
        $this->tags[$tag] = $value;
    }
    /**
     * Getter $tags
     * @return array
     */
    public function getTags()
    {
        return $this->tags;
    }

    /**
     * Getter $unit
     * @return int
     */
    public function getUnit()
    {
        return $this->unit;
    }
    /**
     * Setter $setUnit
     * @param int $unit
     */
    public function setUnit($unit)
    {
        $this->unit = $unit;
    }
}

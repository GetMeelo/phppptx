<?php

/**
 * Transform PPTX to PDF, PPT, ODP
 *
 * @category   Phppptx
 * @package    transform
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */

require_once __DIR__ . '/CreatePptx.php';

abstract class TransformPlugin
{
    /**
     *
     * @access protected
     * @var array
     */
    protected $phppptxconfig;

    /**
     * Construct
     *
     * @access public
     */
    public function __construct()
    {
        $this->phppptxconfig = PhppptxUtilities::parseConfig();
    }

    /**
     * Transform document formats
     *
     * @access public
     * @abstract
     * @param $source
     * @param $target
     * @param array $options
     */
    abstract public function transform($source, $target, $options = array());

    /**
     * Check if the extension is supproted
     *
     * @param string $source
     * @param string $target
     * @param array $supportedExtensionsSource
     * @param array $supportedExtensionsTarget
     * @return array files extensions
     */
    protected function checkSupportedExtension($source, $target, $supportedExtensionsSource, $supportedExtensionsTarget) {
        // get the source file info
        $sourceFileInfo = pathinfo($source);
        $sourceExtension = strtolower($sourceFileInfo['extension']);

        if (!in_array($sourceExtension, $supportedExtensionsSource)) {
            PhppptxLogger::logger('The chosen extension \'' . $sourceExtension . '\' is not supported as source format.', 'fatal');
        }

        // get the target file info
        $targetFileInfo = explode('.', $target);
        $targetExtension = strtolower(array_pop($targetFileInfo));

        if (!in_array($targetExtension, $supportedExtensionsTarget)) {
            PhppptxLogger::logger('The chosen extension \'' . $targetExtension . '\' is not supported as target format.', 'fatal');
        }

        return array('sourceExtension' => $sourceExtension, 'targetExtension' => $targetExtension);
    }
}
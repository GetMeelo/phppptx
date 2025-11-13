<?php

/**
 * Transform documents using MS PowerPoint
 *
 * @category   Phppptx
 * @package    transform
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */

require_once __DIR__ . '/TransformPlugin.php';

class TransformMSPowerPoint extends TransformPlugin
{
    /**
     * Transform documents:
     *     PPTX to PDF, PPT
     *     PPT to PDF, PPTX
     *
     * @access public
     * @param $source string
     * @param $target string
     * @param array $options
     * @throws Exception unsupported file type
     * @throws Exception PHP COM extension is not available
     */
    public function transform($source, $target, $options = array())
    {
        $allowedExtensionsSource = array('ppt', 'pptx');
        $allowedExtensionsTarget = array('ppt', 'pptx', 'pdf');

        $filesExtensions = $this->checkSupportedExtension($source, $target, $allowedExtensionsSource, $allowedExtensionsTarget);

        $code = array(
            'pdf' => new VARIANT(32, VT_I4),
            'ppt' => new VARIANT(1, VT_I4),
            'pptx' => new VARIANT(11, VT_I4),
        );

        // start a PowerPoint instance
        $MSPowerPointInstance = new COM('PowerPoint.application');

        // check that the version of MS PowerPoint is 12 or higher
        if ($MSPowerPointInstance->Version >= 12) {
            // hide MS PowerPoint. This option doesn't work with PowerPoint
            //$MSPowerPointInstance->Visible = 0;

            // open the source document
            $presentation = $MSPowerPointInstance->Presentations->Open($source);

            // save the target document
            $presentation->SaveAs($target, $code[$filesExtensions['targetExtension']]);
        } else {
            PhppptxLogger::logger('The version of PowerPoint should be 12 (PowerPoint 2007) or higher.', 'fatal');
        }
        $MSPowerPointInstance->Quit();

        $MSPowerPointInstance = null;
        unset($MSPowerPointInstance);
    }
}
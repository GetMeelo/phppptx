<?php

/**
 * Transform presentations using native PHP classes
 *
 * @category   Phppptx
 * @package    transform
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */

 require_once __DIR__ . '/TransformPlugin.php';

class TransformNative extends TransformPlugin
{
    /**
     * Transform:
     *     PPTX to HTML
     *
     * @access public
     * @param string $source
     * @param string $target
     * @param array $options
     *  HTML
     *      'htmlPlugin' (TransformNativeHtmlPlugin): plugin to use to do the transformation to HTML. TransformNativeHtmlDefaultPlugin as default
     *      'javaScriptAtTop' (bool) default as false. If true add JS in the head tag
     *      'returnHtmlStructure' (bool) if true return each element of the HTML using an array: css, javascript, metas, presentation. Default as false
     *      'stream' (bool): enable the stream mode. Default as false
     * @throws Exception unsupported file type
     */
    public function transform($source, $target, $options = array())
    {
        $allowedExtensionsSource = array('pptx');
        $allowedExtensionsTarget = array('html');

        $filesExtensions = $this->checkSupportedExtension($source, $target, $allowedExtensionsSource, $allowedExtensionsTarget);

        if ($filesExtensions['sourceExtension'] == 'pptx') {
            if ($filesExtensions['targetExtension'] == 'html') {
                if (!isset($options['htmlPlugin'])) {
                    $options['htmlPlugin'] = new TransformNativeHtmlDefaultPlugin();
                }

                $transform = new TransformNativeHtml($source);
                $html = $transform->transform($options['htmlPlugin'], $options);

                if ((isset($options['stream']) && $options['stream']) || CreatePptx::$streamMode == true) {
                    // stream mode enabled
                    echo $html;
                } else {
                    // stream mode disabled, save the document
                    file_put_contents($target, $html);
                }
            }
        }
    }
}
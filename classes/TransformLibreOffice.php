<?php

/**
 * Transform documents using LibreOffice
 *
 * @category   Phppptx
 * @package    transform
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */

require_once __DIR__ . '/TransformPlugin.php';

class TransformLibreOffice extends TransformPlugin
{
    /**
     * Transform:
     *     PPTX to PDF, PPT, ODP
     *     PPT to PPTX, PDF, ODP
     *     ODP to PPTX, PDF, PPT
     *
     * @access public
     * @param string $source
     * @param string $target
     * @param array $options
     *   'debug' (bool) false (default) or true. Shows debug information about the plugin conversion
     *   'escapeshellarg' (bool) false (default) or true. Applies escapeshellarg to escape source and LibreOffice path strings
     *   'extraOptions' (string) extra parameters to be used when doing the conversion
     *   'homeFolder' (string) set a custom home folder to be used for the conversions
     *   'outdir' (string) set the outdir path. Useful when the PDF output path is not the same than the running script
     *   'path' (string) set the path to LibreOffice. If set, overwrite the path option in phppptxconfig.ini
     * @throws Exception unsupported file type
     */
    public function transform($source, $target, $options = array())
    {
        $allowedExtensionsSource = array('ppt', 'pptx', 'odp');
        $allowedExtensionsTarget = array('ppt', 'pptx', 'pdf', 'odp');

        $filesExtensions = $this->checkSupportedExtension($source, $target, $allowedExtensionsSource, $allowedExtensionsTarget);

        if (!isset($options['debug'])) {
            $options['debug'] = false;
        }
        if (!isset($options['escapeshellarg'])) {
            $options['escapeshellarg'] = false;
        }

        // get the file info
        $sourceFileInfo = pathinfo($source);

        if (isset($options['path'])) {
            $libreOfficePath = $options['path'];
        } else {
            $phppptxconfig = PhppptxUtilities::parseConfig();
            $libreOfficePath = $phppptxconfig['transform']['path'];
        }

        $customHomeFolder = false;
        if (isset($options['homeFolder'])) {
            $currentHomeFolder = getenv("HOME");
            putenv("HOME=" . $options['homeFolder']);
            $customHomeFolder = true;
        } else if (isset($phppptxconfig['transform']['home_folder'])) {
            $currentHomeFolder = getenv("HOME");
            putenv("HOME=" . $phppptxconfig['transform']['home_folder']);
            $customHomeFolder = true;
        }

        $extraOptions = '';
        if (isset($options['extraOptions'])) {
            $extraOptions = $options['extraOptions'];
        }

        // set outputstring for debugging
        $outputDebug = '';
        if (PHP_OS == 'Linux' || PHP_OS == 'Darwin' || PHP_OS == ' FreeBSD') {
            if (!$options['debug']) {
                $outputDebug = ' > /dev/null 2>&1';
            }
        } elseif (substr(PHP_OS, 0, 3) == 'Win' || substr(PHP_OS, 0, 3) == 'WIN') {
            if (!$options['debug']) {
                $outputDebug = ' > nul 2>&1';
            }
        }

        if (isset($options['escapeshellarg']) && $options['escapeshellarg']) {
            // escape file path: empty blank spaces...
            $source = escapeshellarg($source);
            $libreOfficePath = escapeshellarg($libreOfficePath);
        }

        // if the outdir option is set use it as target path, instead use the dir path
        if (isset($options['outdir'])) {
            $outdir = $options['outdir'];
        } else {
            $outdir = $sourceFileInfo['dirname'];
        }

        // call LibreOffice
        passthru($libreOfficePath . ' ' . $extraOptions . ' --invisible --norestore --convert-to ' . $filesExtensions['targetExtension'] . ' ' . $source . ' --outdir ' . $outdir . $outputDebug);

        // get the converted document, this is the name of the source and the extension
        $newDocumentPath = $outdir . '/' . $sourceFileInfo['filename'] . '.' . $filesExtensions['targetExtension'];

        // move the document to the guessed destination
        rename($newDocumentPath, $target);

        // restore the previous HOME value if a custom one has been set
        if ($customHomeFolder && isset($currentHomeFolder)) {
            putenv("HOME=" . $currentHomeFolder);
        }
    }
}
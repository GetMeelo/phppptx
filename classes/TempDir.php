<?php

/**
 * Handle temp directory
 *
 * @category   Phppptx
 * @package    utilities
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class TempDir
{
    /**
     * Return temp dir
     *
     * @access public
     * @return string
     * @static
     */
    public static function getTempDir()
    {
        $phppptxconfig = PhppptxUtilities::parseConfig();

        if (isset($phppptxconfig['settings']['temp_path']) && !empty($phppptxconfig['settings']['temp_path'])) {
            $tempPath = $phppptxconfig['settings']['temp_path'];
        } else {
            $tempPath = sys_get_temp_dir();
        }

        return $tempPath;
    }
}
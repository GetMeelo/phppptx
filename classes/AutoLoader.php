<?php

/**
 * Autoloader
 *
 * @category   Phppptx
 * @package    loader
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class AutoLoader
{
    /**
     * Main tags of relationships XML
     *
     * @access public
     * @static
     */
    public static function load()
    {
        spl_autoload_register(array('AutoLoader', 'autoloadGenericClasses'));
    }

    /**
     * Autoload phppptx
     *
     * @access public
     * @param string $className Class to load
     */
    public static function autoloadGenericClasses($className)
    {
        $pathPhppptx = __DIR__ . '/' . $className . '.php';
        if (file_exists($pathPhppptx)) {
            require_once $pathPhppptx;
        }
    }
}
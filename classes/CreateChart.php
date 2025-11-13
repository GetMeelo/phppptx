<?php

/**
 * Create chart
 *
 * @category   Phppptx
 * @package    elements
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreateChart extends CreateElement
{
    /**
     * Create chart
     *
     * @access public
     * @param string $chart
     * @param array $chartData Chart data
     * @param array $chartStyles
     *      'axPos' (array) position of the axis (r, l, t, b), each value of the array for each position (if a value if null avoids adding it)
     *      'color' (string) (1, 2, 3...) color scheme
     *      'font' (string) Arial, Times New Roman ...
     *      'formatCode' (string) number format
     *      'formatDataLabels' (array)
     *          'rotation' => (int)
     *          'position' => (string) center, insideEnd, insideBase, outsideEnd
     *      'haxLabel' (string) horizontal axis label
     *      'haxLabelDisplay' (string) rotated, vertical, horizontal
     *      'hgrid' (int) 0 (no grid) 1 (only major grid lines - default) 2 (only minor grid lines) 3 (both major and minor grid lines)
     *      'legendOverlay' (bool) if true the legend may overlay the chart
     *      'legendPos' (string) r, l, t, b, none
     *      'majorUnit' (float) bar, col, line charts
     *      'minorUnit' (float) bar, col, line charts
     *      'orientation' (array) orientation of the axis, from min to max (minMax) or max to min (maxMin), each value of the array for each axis (if a value if null avoids adding it)
     *      'scalingMax' (float) scaling max value bar, col, line charts
     *      'scalingMin' (float) scaling min value bar, col, line charts
     *      'showCategory' (bool) shows the categories inside the chart
     *      'showLegendKey' (bool) if true shows the legend values
     *      'showPercent' (bool) if true shows the percent values
     *      'showSeries' (bool) if true shows the series values
     *      'showTable' (bool) if true shows the table of values
     *      'showValue' (bool) if true shows the values inside the chart
     *      'stylesTitle' (array)
     *          'bold' (bool)
     *          'color' (string) FFFFFF, FF0000
     *          'font' (string)  Arial, Times New Roman ...
     *          'fontSize' (int) 8, 9, 10, ... size as drawing content (10 to 400000). 1420 as default
     *          'italic' (bool)
     *          'layout' (array)
     *              'x' (float) 0 < x < 1
     *              'y' (float) 0 < y < 1
     *      'tickLblPos' (mixed) tick label position (nextTo, high, low, none). If string, uses default values. If array, sets a value for each position
     *      'title' (string)
     *      'trendline' (array of trendlines). Compatible with line, bar and col 2D charts
     *          'color' (string) 0000FF
     *          'displayEquation' (bool) display equation on chart
     *          'displayRSquared' (bool) display R-squared value on chart
     *          'intercept' (float) set intercept
     *          'lineStyle' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *          'type' (string) 'exp', 'linear', 'log', 'poly', 'power', 'movingAvg'
     *          'typeOrder' (int) for poly and movingAvg types
     *      'vaxLabel' (string) vertical axis label
     *      'vaxLabelDisplay' (string) rotated, vertical, horizontal
     *      'vgrid'  (int) 0 (no grid) 1 (only major grid lines - default) 2 (only minor grid lines) 3 (both major and minor grid lines)
     *
     *  3D charts:
     *      'perspective' (int) 20, 30...
     *      'rotX' (int) 20, 30...
     *      'rotY' (int) 20, 30...
     *
     *  Bar and column charts:
     *      'gapWidth' (int) gap width
     *      'groupBar' (string) clustered, stacked, percentStacked
     *      'overlap' (int) overlap value
     *
     *  Line charts:
     *      'smooth' (mixed) enable smooth lines, line charts. '0' forces disabling it
     *      'symbol' (string) Line charts: none, dot, plus, square, star, triangle, x, diamond, circle and dash
     *      'symbolSize' (int) the size of the symbols (values 1 to 73)
     *
     *  Pie and doughnut charts:
     *      'explosion' (int) distance between the different values
     *      'holeSize' (int) size of the hole in doughnut type
     *
     *  Theme:
     *  'theme' (array):
     *      'chartArea' (array):
     *          'backgroundColor' (string)
     *      'gridLines' (array):
     *          'capType' (string)
     *          'color' (string): RGB
     *          'dashType' (string)
     *          'width' (int)
     *      'horizontalAxis' (array):
     *          'textBold' (bool)
     *          'textDirection' (string): 'horizontal', 'rotate90', 'rotate270'
     *          'textItalic' (bool)
     *          'textSize' (int): points
     *          'textUnderline' (string): DrawingML values such as 'none', 'sng', 'dash'
     *      'legendArea' (array):
     *          'backgroundColor' (string)
     *          'textBold' (bool)
     *          'textItalic' (bool)
     *          'textSize' (int): points
     *          'textUnderline' (string): DrawingML values such as 'none', 'sng', 'dash'
     *      'plotArea' (array):
     *          'backgroundColor' (string)
     *      'serDataLabels' (array): data labels options (bar, bubble, column, line, ofPie, pie and scatter charts)
     *          'fontStyles' (array):
     *              'bold' (bool)
     *              'color' (string) FFFFFF, FF0000...
     *              'font' (string) Arial, Times New Roman...
     *              'fontSize' (int) size as drawing content (100 to 400000) (100 = 1pt). 1420 as default
     *              'italic' (bool)
     *          'formatCode' (string)
     *          'position (string): center, insideEnd, insideBase, outsideEnd
     *          'showCategory' (int): 0, 1
     *          'showLegendKey' (int): 0, 1
     *          'showPercent' (int): 0, 1
     *          'showSeries' (int): 0, 1
     *          'showValue' (int): 0, 1
     *      'serRgbColors' (array): series colors
     *      'valueDataLabels' (array) data labels options (bar, bubble, column and line charts)
     *          'fontStyles' (array):
     *              'bold' (bool)
     *              'color' (string) FFFFFF, FF0000...
     *              'font' (string) Arial, Times New Roman...
     *              'fontSize' (int) size as drawing content (100 to 400000) (100 = 1pt). 1420 as default
     *              'italic' (bool)
     *          'position' (string): bottom, center, insideEnd, insideBase, left, outsideEnd, right, top, bestFit
     *          'showCategory' (int): 0, 1
     *          'showLegendKey' (int): 0, 1
     *          'showPercent' (int): 0, 1
     *          'showSeries' (int): 0, 1
     *          'showValue' (int): 0, 1
     *          'text' (string) this text hides other automatic values
     *      'valueRgbColors' (array): values colors
     *      'verticalAxis' (array):
     *          'textBold' (bool)
     *          'textDirection' (string): 'horizontal', 'rotate90', 'rotate270'
     *          'textItalic' (bool)
     *          'textSize' (int): points
     *          'textUnderline' (string): DrawingML values such as 'none', 'sng', 'dash'
     * @param array $options
     * @return array
     */
    public function createElementChart($chart, $chartData, $chartStyles = array(), $options = array())
    {
        $newContents = array(
            'chartXml' => '',
            'drawingXml' => '',
        );

        // generate the chart class to be used from the chart type
        $classType = ucwords(str_replace(array('3D', 'Col'), array('', 'Bar'), ucwords($chart)));
        // remove subtype strings
        $classType = str_replace(array('Cylinder', 'Cone', 'Pyramid'), '', $classType);
        $options['type'] = $chart;
        $chartClass = 'CreateChart' . $classType;
        $chartType = new $chartClass();
        $newContents['chartType'] = $chartType;
        // chart content
        $chartXml = $chartType->createChart($chartData, $chartStyles, $options);

        // theme chart
        if (isset($chartStyles['theme']) && is_array($chartStyles['theme']) && count($chartStyles['theme']) > 0) {
            $themeChart = new ThemeCharts();
            $chartXml = $themeChart->theme($chartXml, $chartStyles['theme']);
        }

        $newContents['chartXml'] = $chartXml;

        return $newContents;
    }
}
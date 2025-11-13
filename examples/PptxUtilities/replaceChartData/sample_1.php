<?php
// replace chart data in an existing PPTX

require_once __DIR__ . '/../../../classes/PptxUtilities.php';

$pptx = new PptxUtilities();

$data = array();

// replace charts in the first slide (default)

$data[0] = array(
    'title' => 'A new title',
    'legends' => array(
        'My legend 1',
        'My legend 2',
        'My legend 3',
    ),
    'categories' => array(
        'My cat 1',
        'My cat 2',
        'My cat 3',
        'My cat 4',
    ),
    'values' => array(
        array(25, 10, 5),
        array(20, 5, 4),
        array(15, 0, 3),
        array(10, 15, 2),
    ),
);

$pptx->replaceChartData(__DIR__ . '/../../files/charts.pptx', __DIR__ . '/example_replaceChartData_1.pptx', $data);

// replace charts in the second slide

$data[0] = array(
    'legends' => array(
        'New legend',
    ),
    'categories' => array(
        'cat 1',
        'cat 2',
        'cat 3',
        'cat 4',
    ),
    'values' => array(
        array(25),
        array(20),
        array(15),
        array(10)
    ),
);
$data[1] = array(
    'title' => 'Other title',
    'legends' => array(
        'legend 1',
        'legend 2',
        'legend 3',
    ),
    'categories' => array(
        'other cat 1',
        'other cat 2',
        'other cat 3',
        'other cat 4',
    ),
    'values' => array(
        array(25, 10, 5),
        array(20, 5, 4),
        array(15, 0, 3),
        array(10, 15, 2),
    ),
);

$pptx->replaceChartData(__DIR__ . '/example_replaceChartData_1.pptx', __DIR__ . '/example_replaceChartData_1B.pptx', $data, array('slideNumber' => 2));
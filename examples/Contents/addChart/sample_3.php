<?php
// add a column chart in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx(array('layout' => 'Blank'));

$position = array(
    'coordinateX' => 2400000,
    'coordinateY' => 2500000,
    'sizeX' => 5000000,
    'sizeY' => 2500000,
);
$dataChart = array(
    'legend' => array('Series 1', 'Series 2', 'Series 3'),
    'data' => array(
        array(
            'name' => 'data 1',
            'values' => array(10, 7, 5),
        ),
        array(
            'name' => 'data 2',
            'values' => array(20, 60, 3),
        ),
        array(
            'name' => 'data 3',
            'values' => array(50, 33, 7),
        ),
        array(
            'name' => 'data 4',
            'values' => array(25, 0, 14),
        ),
    ),
);
$stylesChart = array(
    'color' => 2,
    'legendPos' => 'b',
    'legendOverlay' => false,
    'groupBar' => 'stacked',
);
$pptx->addChart('col', $position, $dataChart, $stylesChart);

$pptx->savePptx(__DIR__ . '/example_addChart_3');
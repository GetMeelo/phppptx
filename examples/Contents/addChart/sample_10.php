<?php
// add a surface chart in a PPTX created from scratch

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
            'name' => 'Value1',
            'values' => array(4.3, 2.4, 2),
        ),
        array(
            'name' => 'Value2',
            'values' => array(2.5, 4.4, 2),
        ),
        array(
            'name' => 'Value3',
            'values' => array(3.5, 1.8, 3),
        ),
        array(
            'name' => 'Value4',
            'values' => array(4.5, 2.8, 5),
        ),
        array(
            'name' => 'Value5',
            'values' => array(5, 2, 3),
        ),
    ),
);
$stylesChart = array(
    'color' => 2,
    'legendPos' => 't',
    'legendOverlay' => false,
);
$pptx->addChart('surface', $position, $dataChart, $stylesChart);

$pptx->savePptx(__DIR__ . '/example_addChart_10');
<?php
// add a bar chart in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx(array('layout' => 'Blank'));

$position = array(
    'coordinateX' => 2400000,
    'coordinateY' => 2500000,
    'sizeX' => 5000000,
    'sizeY' => 2500000,
);
$dataChart = array(
    'legend' => array('Legend 1', 'Legend 2', 'Legend 3'),
    'data' => array(
        array(
            'name' => 'data 1',
            'values' => array(10, 20, 5),
        ),
        array(
            'name' => 'data 2',
            'values' => array(20, 60, 3),
        ),
        array(
            'name' => 'data 3',
            'values' => array(50, 33, 7),
        ),
    ),
);
$stylesChart = array(
    'color' => 2,
    'legendOverlay' => false,
    'hgrid' => 1,
    'vgrid' => 2,
);
$pptx->addChart('bar', $position, $dataChart, $stylesChart);

$pptx->savePptx(__DIR__ . '/example_addChart_2');
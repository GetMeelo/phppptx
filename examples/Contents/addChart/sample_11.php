<?php
// add a bubble chart in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx(array('layout' => 'Blank'));

$position = array(
    'coordinateX' => 2400000,
    'coordinateY' => 2500000,
    'sizeX' => 5000000,
    'sizeY' => 2500000,
);
$dataChart = array(
    'legend' => array('', 'values', ''),
    'data' => array(
        array(
            'name' => 'data 1',
            'values' => array(10, 8, 6),
        ),
        array(
            'name' => 'data 2',
            'values' => array(15, 2, 2),
        ),
        array(
            'name' => 'data 3',
            'values' => array(20, 10, 5),
        ),
        array(
            'name' => 'data 4',
            'values' => array(25, 6, 4),
        ),
    ),
);
$stylesChart = array(
    'color' => 28,
    'legendPos' => 't',
    'showTable' => true,
    'showValue' => true,
    'showCategory' => true,
    'hgrid' => '1',
    'vgrid' => '1',
);
$pptx->addChart('bubble', $position, $dataChart, $stylesChart);

$pptx->savePptx(__DIR__ . '/example_addChart_11');
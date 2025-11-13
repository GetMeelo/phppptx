<?php
// add a doughnut chart in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx(array('layout' => 'Blank'));

$position = array(
    'coordinateX' => 2400000,
    'coordinateY' => 2500000,
    'sizeX' => 5000000,
    'sizeY' => 2500000,
);
$dataChart = array(
    'data' => array(
        array(
            'name' => 'data 1',
            'values' => array(20),
        ),
        array(
            'name' => 'data 2',
            'values' => array(20),
        ),
        array(
            'name' => 'data 3',
            'values' => array(50),
        ),
        array(
            'name' => 'data 4',
            'values' => array(25),
        ),
        array(
            'name' => 'data 5',
            'values' => array(5),
        ),
    ),
);
$stylesChart = array(
    'explosion' => 10,
    'holeSize' => 25,
    'color' => 2,
    'legendPos' => 'r',
    'legendOverlay' => true,
    'showPercent' => true,
    'showTable' => true,
);
$pptx->addChart('doughnut', $position, $dataChart, $stylesChart);

$pptx->savePptx(__DIR__ . '/example_addChart_5');
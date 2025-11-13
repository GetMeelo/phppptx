<?php
// add an ofPie chart in a PPTX created from scratch

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
            'values' => array(10),
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
            'values' => array(55),
        ),
        array(
            'name' => 'data 6',
            'values' => array(75),
        ),
        array(
            'name' => 'data 7',
            'values' => array(60),
        ),
        array(
            'name' => 'data 8',
            'values' => array(25),
        ),
    ),
);
$stylesChart = array(
    'title' => 'Pie of pie chart',
    'color' => '26',
    'showPercent' => 1,
    'font' => 'Times New Roman',
    'gapWidth' => 150,
    'secondPieSize' => 75,
    'splitType' => 'val',
    'splitPos' => 30.0,
);
$pptx->addChart('ofPie', $position, $dataChart, $stylesChart);

$pptx->savePptx(__DIR__ . '/example_addChart_13');
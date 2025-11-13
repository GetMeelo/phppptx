<?php
// add a pie chart with a title styled in a PPTX created from scratch

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
    ),
);
$stylesChart = array(
    'title' => 'My title',
    'color' => 2,
    'showPercent' => true,
    'vgrid' => 0,
    'legendPos' => 'r',
    'font' => 'Arial',
    'stylesTitle' => array(
        'bold' => true,
        'color' => 'FF0000',
        'font' => 'Times New Roman',
        'fontSize' => 3600,
        'italic' => true,
    ),
);
$pptx->addChart('pie', $position, $dataChart, $stylesChart);

$pptx->savePptx(__DIR__ . '/example_addChart_7');
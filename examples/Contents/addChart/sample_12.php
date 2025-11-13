<?php
// add a scatter chart in a PPTX created from scratch

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
            'values' => array(10, 0),
        ),
        array(
            'values' => array(17, 2),
        ),
        array(
            'values' => array(18, 4),
        ),
        array(
            'values' => array(25, 6),
        ),
    ),
);
$stylesChart = array(
    'legendPos' => 'r',
    'legendOverlay' => false,
    'haxLabel' => 'hax label',
    'vaxLabel' => 'vax label',
    'haxLabelDisplay' => 'horizontal',
    'vaxLabelDisplay' => 'rotated',
    'hgrid' => 2,
    'vgrid' => 2,
    'symbol' => 'dot',
    'showTable' => true,
    'showValue' => true,
    'showCategory' => true,
);
$pptx->addChart('scatter', $position, $dataChart, $stylesChart);

$pptx->savePptx(__DIR__ . '/example_addChart_12');
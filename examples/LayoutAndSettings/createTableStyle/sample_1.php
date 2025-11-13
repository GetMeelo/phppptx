<?php
// create a table style a apply it to a new table in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx(array('layout' => 'Blank'));

// create a new table style
$tableStyleOptions = array(
    'wholeTbl' => array(
        'backgroundColor' => 'E5F3FD',
        'border' => array(
            'top' => array('color' => 'FF0000'),
            'left' => array('color' => 'FF0000', 'dash' => 'dot'),
            'bottom' => array('color' => 'FF0000', 'width' => 50000),
            'right' => array('color' => 'FF0000', 'dash' => 'dot'),
            'insideV' => array('color' => 'none'),
        ),
    ),
    'firstRow' => array(
        'backgroundColor' => 'FFFF00',
        'border' => array(
            'top' => array('color' => 'none'),
            'left' => array('color' => '00FF00'),
            'bottom' => array('color' => '00FF00'),
            'right' => array('color' => '00FF00'),
        ),
        'bold' => true,
        'italic' => true,
    ),
    'firstCol' => array(
        'backgroundColor' => 'FF0000',
        'border' => array(
            'top' => array('width' => 50000),
            'left' => array('color' => '0000FF'),
            'bottom' => array('color' => '0000FF'),
            'right' => array('color' => '0000FF'),
        ),
    ),
);
$pptx->createTableStyle('MyTableStyle', $tableStyleOptions);

// add a table
$content = array(
    array('Title A', 'Title B', 'Title C'),
    array('Cell 2.1', 'Cell 2.2', 'Cell 2.3'),
    array('Cell 3.1', 'Cell 3.2', 'Cell 3.3'),
    array('Cell 4.1', 'Cell 4.2', 'Cell 4.3'),
);
$position = array(
    'coordinateX' => 4500000,
    'coordinateY' => 2500000,
    'sizeX' => 3500000,
    'sizeY' => 750000,
);
// enable the required styles
$tableStyles = array(
    'firstColumn' => true,
    'headerRow' => true,
    'style' => 'MyTableStyle',
);
// table using a table style and custom styles
$pptx->addTable($content, $position, $tableStyles);

$pptx->savePptx(__DIR__ . '/example_createTableStyle_1');
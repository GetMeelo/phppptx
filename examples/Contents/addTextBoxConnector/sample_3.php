<?php
// add text boxes with text box connectors and fixed positions in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx(array('layout' => 'Blank'));

// add new text boxes
$position = array(
    'coordinateX' => 3500000,
    'coordinateY' => 500000,
    'sizeX' => 750000,
    'sizeY' => 750000,
    'name' => 'My textbox A',
);
$textBoxStyles = array(
    'fill' => array(
        'color' => '03A5FC',
    ),
);
$pptx->addTextBox($position, $textBoxStyles);

$position['coordinateX'] = 1500000;
$position['coordinateY'] = 2500000;
$position['name'] = 'My textbox B';
$pptx->addTextBox($position, $textBoxStyles);

$position['coordinateX'] = 5500000;
$position['coordinateY'] = 2500000;
$position['name'] = 'My textbox C';
$pptx->addTextBox($position, $textBoxStyles);

// add new text contents
$paragraphStyles = array(
    'align' => 'center',
);
$content = array(
    'text' => 'Item A',
);
$pptx->addText($content, array('placeholder' => array('name' => 'My textbox A')), $paragraphStyles);

$content = array(
    'text' => 'Item B',
);
$pptx->addText($content, array('placeholder' => array('name' => 'My textbox B')), $paragraphStyles);

$content = array(
    'text' => 'Item C',
);
$pptx->addText($content, array('placeholder' => array('name' => 'My textbox C')), $paragraphStyles);

// add text box connectors setting fixed positions
$position = array(
    'coordinateX' => 2250000,
    'coordinateY' => 875000,
    'sizeX' => 1250000,
    'sizeY' => 2000000,
);
$connection = array(
    'start' => 'My textbox A',
    'end' => 'My textbox B',
    'positionStart' => 'bottom',
    'positionEnd' => 'top',
);
$options = array(
    'color' => '0000FF',
    'geom' => 'bentConnector3',
    'lineWidth' => 25400,
    'tailEnd' => 'diamond',
    'rotation' => 5400000,
);
$pptx->addTextBoxConnector($position, $connection, $options);

$position = array(
    'coordinateX' => 4250000,
    'coordinateY' => 875000,
    'sizeX' => 1250000,
    'sizeY' => 2000000,
);
$connection = array(
    'start' => 'My textbox A',
    'end' => 'My textbox C',
    'positionStart' => 'bottom',
    'positionEnd' => 'top',
    'flipH' => true,
);
$options = array(
    'color' => '0000FF',
    'geom' => 'bentConnector3',
    'lineWidth' => 25400,
    'tailEnd' => 'diamond',
    'rotation' => 16200000,
);
$pptx->addTextBoxConnector($position, $connection, $options);

$pptx->savePptx(__DIR__ . '/example_addTextBoxConnector_3');
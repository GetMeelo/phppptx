<?php
// add text boxes with text box connectors and custom positions in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx(array('layout' => 'Blank'));

// add new text boxes
$position = array(
    'coordinateX' => 500000,
    'coordinateY' => 3500000,
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

$position['coordinateX'] = 4000000;
$position['coordinateY'] = 2500000;
$position['name'] = 'My textbox B';
$pptx->addTextBox($position, $textBoxStyles);

$position['coordinateX'] = 4000000;
$position['coordinateY'] = 4500000;
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

// add text box connectors setting an automatic position (the position values are not set). These values can also be automatically calculated based on the positions of the text boxes to be connected
$position = array();
$connection = array(
    'start' => 'My textbox A',
    'end' => 'My textbox B',
    'positionStart' => 'left',
    'positionEnd' => 'right',
);
$options = array(
    'color' => '0000FF',
    'geom' => 'bentConnector5',
    'lineWidth' => 25400,
    'shapeGuide' => array(
        array(
            'fmla' => 'val -5268',
            'guide' => 'adj1',
        ),
        array(
            'fmla' => 'val 50000',
            'guide' => 'adj2',
        ),
        array(
            'fmla' => 'val 105268',
            'guide' => 'adj3',
        ),
    ),
    'tailEnd' => 'diamond',
);
$pptx->addTextBoxConnector($position, $connection, $options);

$connection = array(
    'start' => 'My textbox A',
    'end' => 'My textbox C',
);
$options = array(
    'color' => '0000FF',
    'geom' => 'bentConnector3',
    'lineWidth' => 25400,
    'tailEnd' => 'diamond',
);
$pptx->addTextBoxConnector($position, $connection, $options);

$pptx->savePptx(__DIR__ . '/example_addTextBoxConnector_2');
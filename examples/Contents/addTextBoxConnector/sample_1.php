<?php
// add text boxes with text box connectors in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx(array('layout' => 'Blank'));

// add new text boxes
$position = array(
    'coordinateX' => 770000,
    'coordinateY' => 500000,
    'sizeX' => 1200000,
    'sizeY' => 750000,
    'name' => 'Custom textbox 1',
);
$textBoxStyles = array(
    'autofit' => 'noautofit',
    'fill' => array(
        'color' => '99B369',
    ),
);
$pptx->addTextBox($position, $textBoxStyles);

$position = array(
    'coordinateX' => 5400000,
    'coordinateY' => 500000,
    'sizeX' => 1200000,
    'sizeY' => 750000,
    'name' => 'Custom textbox 2',
);
$textBoxStyles = array(
    'autofit' => 'noautofit',
    'fill' => array(
        'color' => '99B369',
    ),
);
$pptx->addTextBox($position, $textBoxStyles);

// add text contents in the text boxes
$content = array(
    'text' => 'Textbox 1.',
    'bold' => true,
    'underline' => 'single',
);
$pptx->addText($content, array('placeholder' => array('name' => 'Custom textbox 1')));

$content = array(
    'text' => 'Textbox 2.',
    'bold' => true,
    'underline' => 'single',
);
$pptx->addText($content, array('placeholder' => array('name' => 'Custom textbox 2')));

// the getActiveSlideInformation method can be used to return text box ids and names to be connected
$activeSlideInformation = $pptx->getActiveSlideInformation();

// add a text box connector setting a fixed position
$position = array(
    'coordinateX' => 1970000,
    'coordinateY' => 875000,
    'sizeX' => 3430000,
    'sizeY' => 0,
);
// add the connector using text box IDs. Text box names can be used to connect text boxes
$connection = array(
    'start' => (int)$activeSlideInformation['placeholders'][0]['id'],
    'end' => (int)$activeSlideInformation['placeholders'][1]['id'],
);
$options = array(
    'color' => 'FF0000',
    'dash' => 'sysDash',
    'lineWidth' => 25400,
);
$pptx->addTextBoxConnector($position, $connection, $options);

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
);
$options = array(
    'color' => '0000FF',
    'geom' => 'bentConnector3',
    'lineWidth' => 25400,
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

$pptx->savePptx(__DIR__ . '/example_addTextBoxConnector_1');
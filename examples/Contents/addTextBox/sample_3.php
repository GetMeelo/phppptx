<?php
// add text boxes with styles in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');

// add the text boxed in the second slide. The first slide has the position 0
$pptx->setActiveSlide(array('position' => 1));

$position = array(
    'coordinateX' => 6470000,
    'coordinateY' => 600000,
    'sizeX' => 5000000,
    'sizeY' => 450000,
    'name' => 'Textbox custom A',
);
$textBoxStyles = array(
    'border' => array(
        'color' => '000000',
    ),
);
$pptx->addTextBox($position, $textBoxStyles);

$position = array(
    'coordinateX' => 1770000,
    'coordinateY' => 3507000,
    'sizeX' => 4000000,
    'sizeY' => 2350000,
    'name' => 'Textbox custom A',
    'order' => 0,
);
$textBoxStyles = array(
    'autofit' => 'noautofit',
    'fill' => array(
        'image' => __DIR__ . '/../../files/image.png',
    ),
);
$pptx->addTextBox($position, $textBoxStyles);

$pptx->savePptx(__DIR__ . '/example_addTextBox_3');
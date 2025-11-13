<?php
// add shapes in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');

// change the active slide
$pptx->setActiveSlide(array('position' => -1));

$position = array(
    'coordinateX' => 380000,
    'coordinateY' => 760000,
    'sizeX' => 500000,
    'sizeY' => 450000,
);
$options = array(
    'fillColor' => '#0000FF',
);
$pptx->addShape('rightArrow', $position, $options);

$pptx->savePptx(__DIR__ . '/example_addShape_2');
<?php
// add tables applying styles in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');

$pptx->setActiveSlide(array('position' => 1));

$content = array(
    array('Title A', 'Title B'),
    array('Cell 2.1', 'Cell 2.2'),
    array('Cell 3.1', 'Cell 3.2'),
);
$position = array(
    'coordinateX' => 7500000,
    'coordinateY' => 2500000,
    'sizeX' => 2500000,
    'sizeY' => 750000,
);
$tableStyles = array(
    'bandedRows' => true,
    'headerRow' => true,
    'style' => 'Medium Style 2 - Accent 1',
);
// table using a table style and custom styles
$pptx->addTable($content, $position, $tableStyles);

$pptx->savePptx(__DIR__ . '/example_addTable_5');
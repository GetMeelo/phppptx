<?php
// add an SVG image file

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$position = array(
    'coordinateX' => 1400000,
    'coordinateY' => 4500000,
    'sizeX' => 2200000,
    'sizeY' => 2000000,
);

$pptx->addSvg(__DIR__ . '/../../files/image.svg', $position);

$pptx->savePptx(__DIR__ . '/example_addSvg_1');
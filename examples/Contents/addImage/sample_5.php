<?php
// add an image using an image resource

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$imageResource = imagecreatefromjpeg(__DIR__ . '/../../files/image.jpg');
$position = array(
    'coordinateX' => 1400000,
    'coordinateY' => 4500000,
);
$pptx->addImage($imageResource, $position);

$pptx->savePptx(__DIR__ . '/example_addImage_5');
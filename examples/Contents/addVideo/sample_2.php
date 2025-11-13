<?php
// add a video with a custom image in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$position = array(
    'coordinateX' => 1400000,
    'coordinateY' => 4500000,
    'sizeX' => 2000000,
    'sizeY' => 2000000,
);
$videoStyles = array(
    'image' => array(
        'image' => __DIR__ . '/../../files/image.png',
    )
);
$pptx->addVideo(__DIR__ . '/../../files/video.mp4', $position, $videoStyles);

$pptx->savePptx(__DIR__ . '/example_addVideo_2');
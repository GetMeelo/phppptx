<?php
// add an audio with a custom image in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$position = array(
    'coordinateX' => 1400000,
    'coordinateY' => 4500000,
    'sizeX' => 2000000,
    'sizeY' => 2000000,
);
$audioStyles = array(
    'image' => array(
        'image' => __DIR__ . '/../../files/image.png',
    )
);
$pptx->addAudio(__DIR__ . '/../../files/audio.mp3', $position, $audioStyles);

$pptx->savePptx(__DIR__ . '/example_addAudio_2');
<?php
// move audios, images and videos

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// change the active slide position
$pptx->setActiveSlide(array('position' => 2));

// move the images in the active slide to the second slide
$referenceNode = array(
    'type' => 'image',
);
$referenceNodeTo = array(
    'type' => 'slide',
    'occurrence' => 2,
);
$pptx->moveElement($referenceNode, $referenceNodeTo);

// change the active slide position
$pptx->setActiveSlide(array('position' => 4));

// move the audios in the active slide to the first slide
$referenceNode = array(
    'type' => 'audio',
);
$referenceNodeTo = array(
    'type' => 'slide',
    'occurrence' => 1,
);
$pptx->moveElement($referenceNode, $referenceNodeTo);

// move the audios in the active slide to the third slide
$referenceNode = array(
    'type' => 'video',
);
$referenceNodeTo = array(
    'type' => 'slide',
    'occurrence' => 3,
);
$pptx->moveElement($referenceNode, $referenceNodeTo);

$pptx->savePptx(__DIR__ . '/example_moveElement_4');
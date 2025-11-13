<?php
// move a shape

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// move the second shape in the active slide to the fourth slide
$referenceNode = array(
    'type' => 'shape',
    'occurrence' => 2,
);
$referenceNodeTo = array(
    'type' => 'slide',
    'occurrence' => 4,
);
$pptx->moveElement($referenceNode, $referenceNodeTo);

// change the active slide position
$pptx->setActiveSlide(array('position' => 2));

$pptx->savePptx(__DIR__ . '/example_moveElement_2');
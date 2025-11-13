<?php
// move a table

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// change the active slide position
$pptx->setActiveSlide(array('position' => 3));

// move the second table in the active slide to the second slide
$referenceNode = array(
    'type' => 'table',
    'occurrence' => 2,
);
$referenceNodeTo = array(
    'type' => 'slide',
    'occurrence' => 2,
);
$pptx->moveElement($referenceNode, $referenceNodeTo);

$pptx->savePptx(__DIR__ . '/example_moveElement_3');
<?php
// clone an image

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// change the active slide position
$pptx->setActiveSlide(array('position' => 2));

$referenceNode = array(
    'type' => 'image',
    'occurrence' => 1,
);
$pptx->cloneElement($referenceNode);

$pptx->savePptx(__DIR__ . '/example_cloneElement_3');
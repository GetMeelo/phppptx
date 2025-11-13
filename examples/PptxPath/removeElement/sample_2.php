<?php
// remove images

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// change the active slide position
$pptx->setActiveSlide(array('position' => 2));

$referenceNode = array(
    'type' => 'image',
);
$pptx->removeElement($referenceNode);

$pptx->savePptx(__DIR__ . '/example_removeElement_2');
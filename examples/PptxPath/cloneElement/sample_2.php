<?php
// clone paragraphs

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// clone the first paragraph
$referenceNode = array(
    'type' => 'paragraph',
    'occurrence' => 1,
);
$pptx->cloneElement($referenceNode);

// change the active slide position
$pptx->setActiveSlide(array('position' => 1));

// clone the first paragraph that contains 'features' text
$referenceNode = array(
    'type' => 'paragraph',
    'contains' => 'features',
);
$pptx->cloneElement($referenceNode);

$pptx->savePptx(__DIR__ . '/example_cloneElement_2');
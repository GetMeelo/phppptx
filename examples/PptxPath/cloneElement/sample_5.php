<?php
// clone shapes

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// clone the first shape
$referenceNode = array(
    'type' => 'shape',
    'occurrence' => 1,
);
$pptx->cloneElement($referenceNode);

// change the active slide position
$pptx->setActiveSlide(array('position' => 1));

// clone the shapes that include "Content Placeholder 2" as name
$referenceNode = array(
    'type' => 'shape',
    'attributes' => array(
        'p:cNvPr' => array(
            'name' => 'Content Placeholder 2',
        ),
    ),
);
$pptx->cloneElement($referenceNode);

$pptx->savePptx(__DIR__ . '/example_cloneElement_5');
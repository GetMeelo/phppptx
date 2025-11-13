<?php
// clone slides

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// clone slide 2
$referenceNode = array(
    'type' => 'slide',
    'occurrence' => 2,
);
$pptx->cloneElement($referenceNode);

// clone slide 4
$referenceNode = array(
    'type' => 'slide',
    'occurrence' => 4,
);
$pptx->cloneElement($referenceNode);

$pptx->savePptx(__DIR__ . '/example_cloneElement_1');
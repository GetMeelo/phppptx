<?php
// move slides

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// move slide 2 after slide 4
$referenceNode = array(
    'type' => 'slide',
    'occurrence' => 2,
);
$referenceNodeTo = array(
    'type' => 'slide',
    'occurrence' => 4,
);
$pptx->moveElement($referenceNode, $referenceNodeTo);

// set the last slide as the first slide
$referenceNode = array(
    'type' => 'slide',
    'occurrence' => 'last()',
);
$referenceNodeTo = array(
    'type' => 'slide',
    'occurrence' => 'first()',
);
$pptx->moveElement($referenceNode, $referenceNodeTo);

$pptx->savePptx(__DIR__ . '/example_moveElement_1');
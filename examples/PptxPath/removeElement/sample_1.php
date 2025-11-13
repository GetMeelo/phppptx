<?php
// remove a slide

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

$referenceNode = array(
    'type' => 'slide',
    'occurrence' => 2,
);
$pptx->removeElement($referenceNode);

$pptx->savePptx(__DIR__ . '/example_removeElement_1');
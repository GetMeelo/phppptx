<?php
// remove table contents

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// change the active slide position
$pptx->setActiveSlide(array('position' => 3));

// remove the first table
$referenceNode = array(
    'type' => 'table',
    'occurrence' => 1,
);
$pptx->removeElement($referenceNode);

// remove the row that contains '$DESCRIPTION$' text
$referenceNode = array(
    'type' => 'table-row',
    'contains' => '$DESCRIPTION$',
);
$pptx->removeElement($referenceNode);

$pptx->savePptx(__DIR__ . '/example_removeElement_5');
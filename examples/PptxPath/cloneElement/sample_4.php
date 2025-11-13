<?php
// clone a table and a table row

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// change the active slide position
$pptx->setActiveSlide(array('position' => 3));

// clone the first table
$referenceNode = array(
    'type' => 'table',
    'occurrence' => 1,
);
$pptx->cloneElement($referenceNode);

// clone the row that contains '$DESCRIPTION$' text
$referenceNode = array(
    'type' => 'table-row',
    'contains' => '$DESCRIPTION$',
);
$pptx->cloneElement($referenceNode);

$pptx->savePptx(__DIR__ . '/example_cloneElement_4');
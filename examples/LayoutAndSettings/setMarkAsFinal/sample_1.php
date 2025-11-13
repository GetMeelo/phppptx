<?php
// set mark as final to the document to prevent changing it in a PPTX created from scratch. Premium licenses include crypto and sign features to get better protection

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$pptx->setMarkAsFinal();

$pptx->savePptx(__DIR__ . '/example_setMarkAsFinal_1');
<?php
// add a new section in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();
$pptx->addSection();

$pptx->savePptx(__DIR__ . '/example_addSection_1');
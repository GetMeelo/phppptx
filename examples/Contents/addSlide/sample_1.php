<?php
// add a new slide in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();
$pptx->addSlide();

$pptx->savePptx(__DIR__ . '/example_addSlide_1');
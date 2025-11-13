<?php
// add a macro in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();
$pptx->addMacroFromPptx(__DIR__ . '/../../files/sample_macro.pptm');

$pptx->savePptx(__DIR__ . '/example_addMacroFromPptx_1.pptm');
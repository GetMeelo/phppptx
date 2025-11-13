<?php
// remove text variables (placeholders)

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// remove variables
$pptx->removeVariableText(array('VAR_TITLE'));

$pptx->savePptx(__DIR__ . '/example_removeVariableText_1');
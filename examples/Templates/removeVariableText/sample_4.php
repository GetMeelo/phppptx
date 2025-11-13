<?php
// remove text variables (placeholders) using the block type removal

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// remove variables
$pptx->removeVariableText(array('VAR_TITLE'), array('type' => 'block'));

$pptx->savePptx(__DIR__ . '/example_removeVariableText_4');
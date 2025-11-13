<?php
// replace text variables (placeholders) with new text using the block type removal

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// replace variables
$pptx->replaceVariableText(array('VAR_TITLE' => 'phppptx'), array('type' => 'block'));

$pptx->savePptx(__DIR__ . '/example_replaceVariableText_5');
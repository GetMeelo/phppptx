<?php
// change the symbol used to wrap variables (placehoders) and replace text variables (placeholder)

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_symbols.pptx');
$pptx->setTemplateSymbol('${', '}');

// replace variables
$pptx->replaceVariableText(array('VAR_TITLE' => 'phppptx'));

$pptx->savePptx(__DIR__ . '/example_setTemplateSymbol_1');
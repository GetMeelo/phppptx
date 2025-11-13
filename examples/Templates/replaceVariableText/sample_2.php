<?php
// replace text variables (placeholders) with new text contents using ${} to wrap placeholders

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_symbols.pptx');
$pptx->setTemplateSymbol('${', '}');

// replace variables
$pptx->replaceVariableText(array('VAR_TITLE' => 'phppptx'));

$pptx->savePptx(__DIR__ . '/example_replaceVariableText_2');
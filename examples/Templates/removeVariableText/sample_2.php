<?php
// remove text variables (placeholders) using ${} to wrap placeholders

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_symbols.pptx');
$pptx->setTemplateSymbol('${', '}');

// remove variables
$pptx->removeVariableText(array('VAR_TITLE'));

$pptx->savePptx(__DIR__ . '/example_removeVariableText_2');
<?php
// return the variables (placeholders) using ${} to wrap placeholders

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_symbols.pptx');
$pptx->setTemplateSymbol('${', '}');

print_r($pptx->getTemplateVariables());
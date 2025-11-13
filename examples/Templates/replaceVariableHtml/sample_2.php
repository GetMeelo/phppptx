<?php
// replace text variables (placeholders) with HTML using ${} to wrap placeholders doing block type replacements

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_symbols.pptx');
$pptx->setTemplateSymbol('${', '}');

$html = '<p style="font-size: 48px;"><b>php</b><em style="color:#C060E0;">pptx</em></p>';

// replace variables
$pptx->replaceVariableHtml(array('VAR_TITLE' => $html));

$pptx->savePptx(__DIR__ . '/example_replaceVariableHtml_2');
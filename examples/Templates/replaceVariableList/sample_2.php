<?php
// replace list variables (placeholders) using ${} to wrap placeholders

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_symbols.pptx');
$pptx->setTemplateSymbol('${', '}');

$items = array('First item', 'Second item', array('Subitem A', 'Subitem B', 'Subitem C', array('Subitem C.1', 'Subitem C.2'), 'Subitem D'), 'Third item');

$pptx->replaceVariableList('VAR_LIST', $items);

$pptx->savePptx(__DIR__ . '/example_replaceVariableList_2');
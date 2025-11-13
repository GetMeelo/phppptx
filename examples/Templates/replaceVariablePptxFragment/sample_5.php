<?php
// replace text variables (placeholders) with PptxFragments using ${} to wrap placeholders doing inline type replacements

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_symbols.pptx');
$pptx->setTemplateSymbol('${', '}');

$contentA = array(
    'text' => 'phppptx',
    'bold' => true,
    'underline' => 'single',
);
$textFragmentA = new PptxFragment();
$textFragmentA->addText($contentA, array());

// replace variables
$pptx->replaceVariablePptxFragment(array('VAR_TITLE' => $textFragmentA), array('type' => 'inline'));

$pptx->savePptx(__DIR__ . '/example_replaceVariablePptxFragment_5');
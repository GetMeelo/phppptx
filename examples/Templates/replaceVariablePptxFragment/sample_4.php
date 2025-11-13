<?php
// replace text variables (placeholders) with PptxFragments doing inline type replacements

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

$contentA = array(
    'text' => 'phppptx',
    'bold' => true,
    'underline' => 'single',
);
$textFragmentA = new PptxFragment();
$textFragmentA->addText($contentA, array());

// replace variables
$pptx->replaceVariablePptxFragment(array('VAR_TITLE' => $textFragmentA), array('type' => 'inline'));

$pptx->savePptx(__DIR__ . '/example_replaceVariablePptxFragment_4');
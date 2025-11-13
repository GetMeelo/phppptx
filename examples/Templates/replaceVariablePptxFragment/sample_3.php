<?php
// replace text variables (placeholders) with PptxFragments in the internal active slide doing block type replacements

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

$contentA = array(
    'text' => 'phppptx',
    'color' => 'FFFFFF',
);
$textFragmentA = new PptxFragment();
$textFragmentA->addText($contentA, array());

// replace variables in the internal active slide
$pptx->replaceVariablePptxFragment(array('VAR_TITLE' => $textFragmentA), array('activeSlide' => true));

// change the internal active slide
$pptx->setActiveSlide(array('position' => 1));

$contentB = array(
    'text' => 'Another value',
    'bold' => true,
    'color' => '000000',
);
$paragraphStyles = array(
    'align' => 'center',
);
$textFragmentB = new PptxFragment();
$textFragmentB->addText($contentB, array(), $paragraphStyles);

// replace variables in the internal active slide
$pptx->replaceVariablePptxFragment(array('VAR_TITLE' => $textFragmentB), array('activeSlide' => true));

$pptx->savePptx(__DIR__ . '/example_replaceVariablePptxFragment_3');
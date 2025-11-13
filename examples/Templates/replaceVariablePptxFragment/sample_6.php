<?php
// replace text variables (placeholders) with PptxFragments doing block replacements in notesSlides, slides, slideLayouts and slideMasters targets

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_multi.pptx');

$content = array(
    'text' => 'phppptx',
    'bold' => true,
    'underline' => 'single',
);
$textFragment = new PptxFragment();
$textFragment->addText($content, array());
$pptx->replaceVariablePptxFragment(array('VAR_TITLE' => $textFragment));

$content = array(
    'text' => 'New content in layout',
    'bold' => true,
    'underline' => 'single',
);
$textFragment = new PptxFragment();
$textFragment->addText($content, array());
$pptx->replaceVariablePptxFragment(array('VAR_TITLE' => $textFragment), array('target' => 'slideLayouts'));

$content = array(
    'text' => 'New content in slide master',
    'bold' => true,
    'underline' => 'single',
);
$textFragment = new PptxFragment();
$textFragment->addText($content, array());
$pptx->replaceVariablePptxFragment(array('VAR_MASTER' => $textFragment), array('target' => 'slideMasters'));


$content = array(
    'text' => 'New content in comment',
    'bold' => true,
    'underline' => 'single',
);
$textFragment = new PptxFragment();
$textFragment->addText($content, array());
$pptx->replaceVariablePptxFragment(array('VAR_NOTE_1' => $textFragment), array('target' => 'notesSlides'));

$pptx->savePptx(__DIR__ . '/example_replaceVariablePptxFragment_6');
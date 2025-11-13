<?php
// replace text variables (placeholders) with new text in the internal active slide

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// replace variables in the internal active slide
$pptx->replaceVariableText(array('VAR_TITLE' => 'phppptx'), array('activeSlide' => true));

// change the internal active slide
$pptx->setActiveSlide(array('position' => 1));

// replace variables in the internal active slide
$pptx->replaceVariableText(array('VAR_TITLE' => 'Another value'), array('activeSlide' => true));

$pptx->savePptx(__DIR__ . '/example_replaceVariableText_3');
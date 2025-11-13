<?php
// remove text variables (placeholders) in notesSlides, slides, slideLayouts and slideMasters targets

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_multi.pptx');

// remove variables
$pptx->removeVariableText(array('VAR_TITLE', 'VAR_TITLE_2'));
$pptx->removeVariableText(array('VAR_TITLE'), array('target' => 'slideLayouts'));
$pptx->removeVariableText(array('VAR_MASTER'), array('target' => 'slideMasters'));
$pptx->removeVariableText(array('VAR_NOTE_1', 'VAR_NOTE_2', 'VAR_NOTE_3'), array('target' => 'notesSlides'));

$pptx->savePptx(__DIR__ . '/example_removeVariableText_3');
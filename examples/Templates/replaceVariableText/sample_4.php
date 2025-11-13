<?php
// replace text variables (placeholders) with new text in notesSlides, slides, slideLayouts and slideMasters targets

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_multi.pptx');

// replace variables in slides target
$pptx->replaceVariableText(array('VAR_TITLE' => 'phppptx', 'VAR_SUBTITLE' => 'phppptx'));

// replace variables in slideLayouts target
$pptx->replaceVariableText(array('VAR_TITLE' => 'title content', 'VAR_SUBTITLE' => 'subtitle content'), array('target' => 'slideLayouts'));

// replace variables in slideMasters target
$pptx->replaceVariableText(array('VAR_MASTER' => 'master'), array('target' => 'slideMasters'));

// replace variables in notesSlides target
$pptx->replaceVariableText(array('VAR_NOTE_1' => 'A new note.', 'VAR_NOTE_2' => 'Note 2', 'VAR_NOTE_3' => 'Note 3'), array('target' => 'notesSlides'));

$pptx->savePptx(__DIR__ . '/example_replaceVariableText_4');
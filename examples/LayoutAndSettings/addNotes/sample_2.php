<?php
// add notes in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_multi.pptx');

// add notes in the first slide. By default, new notes are added after existing notes
$pptx->addNotes('A new note.');

// add notes in the second slide. Replace existing notes
$pptx->setActiveSlide(array('position' => 1));
$pptx->addNotes('A new note in the second slide.', array(), array('insertMode' => 'replace'));

// add notes in the third slide
$pptx->setActiveSlide(array('position' => 2));
$pptx->addNotes('A new note in the third slide.');

$pptx->savePptx(__DIR__ . '/example_addNotes_2');
<?php
// remove notes in the internal active slide in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_multi.pptx');

// set the first slide as the internal active slide
$pptx->setActiveSlide(array('position' => 0));
// remove notes from the internal active slide
$pptx->removeNotesSlide();

// set the second slide as the internal active slide
$pptx->setActiveSlide(array('position' => 1));
// remove notes from the internal active slide
$pptx->removeNotesSlide();

$pptx->savePptx(__DIR__ . '/example_removeNotesSlide_1');
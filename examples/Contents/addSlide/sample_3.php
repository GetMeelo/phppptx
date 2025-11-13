<?php
// add new slides setting layouts in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');
$pptx->addSlide(array('layout' => 'Title Slide'));
$pptx->addSlide(array('layout' => 'Title and Content'));

$pptx->savePptx(__DIR__ . '/example_addSlide_3');
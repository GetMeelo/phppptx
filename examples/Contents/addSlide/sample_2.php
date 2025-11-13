<?php
// add new slides setting layouts in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();
$pptx->addSlide(array('layout' => 'Title and Content'));
$pptx->addSlide(array('layout' => 'Comparison'));

$pptx->savePptx(__DIR__ . '/example_addSlide_2');
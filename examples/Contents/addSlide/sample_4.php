<?php
// new PPTX created from scratch setting a custom layout to be used in the first slide.
// The default layout in a PPTX created from scratch is 'Title Slide'.

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx(array('layout' => 'Title and Content'));
$pptx->addSlide(array('layout' => 'Two Content'));

$pptx->savePptx(__DIR__ . '/example_addSlide_4');
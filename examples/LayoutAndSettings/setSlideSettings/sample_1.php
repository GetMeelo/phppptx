<?php
// modify the layout of the second and third slides in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');

// the first slide has the position 0. Set the second slide as the internal active slide
$pptx->setActiveSlide(array('position' => 1));
$pptx->setSlideSettings(array('layout' => 'Title Slide'));

// the first slide has the position 0. Set the third slide as the internal active slide
$pptx->setActiveSlide(array('position' => 2));
$pptx->setSlideSettings(array('layout' => 'Blank'));

$pptx->savePptx(__DIR__ . '/example_setSlideSettings_1');
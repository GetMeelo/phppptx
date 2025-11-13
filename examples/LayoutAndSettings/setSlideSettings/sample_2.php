<?php
// hide a slide in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');

// the first slide has the position 0. Set the second slide as the internal active slide
$pptx->setActiveSlide(array('position' => 1));

// hide the slide
$pptx->setSlideSettings(array('show' => false));

$pptx->savePptx(__DIR__ . '/example_setSlideSettings_2');
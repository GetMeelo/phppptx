<?php
// get the internal active slide information in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');

$activeSlideInformation = $pptx->getActiveSlideInformation();
var_dump($activeSlideInformation);

// the first slide has the position 0
$pptx->setActiveSlide(array('position' => 1));

$activeSlideInformation = $pptx->getActiveSlideInformation();
var_dump($activeSlideInformation);
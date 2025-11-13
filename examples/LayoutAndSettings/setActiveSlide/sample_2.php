<?php
// set the internal active slide in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');
// the default internal active slide is the last one
$activeSlide = $pptx->getActiveSlide();
var_dump($activeSlide);

// the first slide has the position 0
$pptx->setActiveSlide(array('position' => 0));

$activeSlide = $pptx->getActiveSlide();
var_dump($activeSlide);

// the last slide can be set as the internal active slide using its exact position or -1
$pptx->setActiveSlide(array('position' => 2));

$activeSlide = $pptx->getActiveSlide();
var_dump($activeSlide);

$pptx->setActiveSlide(array('position' => -1));

$activeSlide = $pptx->getActiveSlide();
var_dump($activeSlide);

// a non-existing position returns an Exception
//$pptx->setActiveSlide(array('position' => 3));
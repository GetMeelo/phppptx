<?php
// set slide background colors in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');
$pptx->setActiveSlide(array('position' => 0));
$pptx->setBackgroundColor('00B0F0');

$pptx->savePptx(__DIR__ . '/example_setBackgroundColor_2');
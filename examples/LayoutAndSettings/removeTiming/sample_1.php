<?php
// remove timing element

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

$pptx->removeTiming();

$pptx->setActiveSlide(array('position' => 4));

$pptx->removeTiming();

$pptx->savePptx(__DIR__ . '/example_removeTiming_1');
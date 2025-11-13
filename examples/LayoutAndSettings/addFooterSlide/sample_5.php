<?php
// add footer slide in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');

$pptx->addFooterSlide('dateAndTime', array('applyToAll' => true));

$pptx->savePptx(__DIR__ . '/example_addFooterslide_5');
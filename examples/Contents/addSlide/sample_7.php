<?php
// add a new slide in a PPTX created from scratch. Do not clean paragraph contents from the layout

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();
$pptx->addSlide(array('cleanLayoutParagraphContents' => false));

$pptx->savePptx(__DIR__ . '/example_addSlide_7');
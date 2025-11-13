<?php
// set rtl settings in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();
$pptx->setPresentationSettings(array('rtl' => true));
$pptx->setRtl();

$pptx->savePptx(__DIR__ . '/example_setRtl_1');
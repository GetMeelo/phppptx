<?php
// modify the size of the presentation using a preset size in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');
$pptx->setPresentationSettings(array('type' => 'A3'));

$pptx->savePptx(__DIR__ . '/example_setPresentationSettings_3');
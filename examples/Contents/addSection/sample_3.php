<?php
// add new sections in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_sections.pptx');
$pptx->addSection();
$pptx->addSection(array('name' => 'My section'));

$pptx->savePptx(__DIR__ . '/example_addSection_3');
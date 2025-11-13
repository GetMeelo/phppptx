<?php
// process a template

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');
$pptx->processTemplate();

$pptx->savePptx(__DIR__ . '/example_processTemplate_1');
<?php
// return the variables (placeholders)

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

print_r($pptx->getTemplateVariablesType());
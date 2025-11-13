<?php
// return the variables (placeholders) in multiple targets

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_multi.pptx');

print_r($pptx->getTemplateVariables());
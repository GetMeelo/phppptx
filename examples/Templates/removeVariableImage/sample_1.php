<?php
// remove image variables (placeholders). The placeholders have been added to the alt text

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// remove images in slides target
$pptx->removeVariableImage(array('VAR_IMAGE_PHPDOCX', 'VAR_IMAGE_PHPXLSX'));

$pptx->savePptx(__DIR__ . '/example_removeVariableImage_1');
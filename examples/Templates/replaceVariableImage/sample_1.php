<?php
// replace image variables (placeholders). The placeholders have been added to the alt text

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// replace images in slides target
$pptx->replaceVariableImage(
    array(
        'VAR_IMAGE_PHPDOCX' => __DIR__ . '/../../files/imageP1.png',
        'VAR_IMAGE_PHPXLSX' => __DIR__ . '/../../files/imageP2.png',
    )
);

$pptx->savePptx(__DIR__ . '/example_replaceVariableImage_1');
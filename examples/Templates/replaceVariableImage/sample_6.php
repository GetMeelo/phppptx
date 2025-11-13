<?php
// replace image variables (placeholders) setting custom sizes and descr values. The placeholders have been added to the alt text

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// replace images in slides target
$options = array(
    'descr' => 'New descr',
    'sizeX' => 3500000,
    'sizeY' => 2000000,
);
$pptx->replaceVariableImage(
    array(
        'VAR_IMAGE_PHPDOCX' => __DIR__ . '/../../files/imageP1.png',
    ),
    $options
);

$pptx->savePptx(__DIR__ . '/example_replaceVariableImage_6');
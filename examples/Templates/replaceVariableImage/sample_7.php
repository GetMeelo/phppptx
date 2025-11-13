<?php
// replace image variables (placeholders) in slides and slideLayouts targets. The placeholders have been added to the alt text

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_multi.pptx');

// replace images in slides target
$pptx->replaceVariableImage(
    array(
        'VAR_IMG_SLIDE2' => __DIR__ . '/../../files/imageP1.png',
    )
);

// replace images in slideLayouts target
$pptx->replaceVariableImage(
    array(
        'VAR_IMAGE' => __DIR__ . '/../../files/imageP2.png',
    ),
    array('target' => 'slideLayouts')
);

$pptx->savePptx(__DIR__ . '/example_replaceVariableImage_7');
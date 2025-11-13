<?php
// replace image variables (placeholders) using ${} to wrap placeholders. The placeholder has been added to the alt text content

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_symbols.pptx');
$pptx->setTemplateSymbol('${', '}');

// replace images in slides target
$pptx->replaceVariableImage(
    array(
        'VAR_IMAGE_PHPDOCX' => __DIR__ . '/../../files/imageP1.png',
        'VAR_IMAGE_PHPXLSX' => __DIR__ . '/../../files/imageP2.png',
    )
);

$pptx->savePptx(__DIR__ . '/example_replaceVariableImage_3');
<?php
// replace image variables (placeholders) using stream sources. The placeholders have been added to the alt text content

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

$pptx->replaceVariableImage(
    array(
        'VAR_IMAGE_PHPDOCX' => 'https://www.phpdocx.com/img/logo_badge.png',
        'VAR_IMAGE_PHPXLSX' => 'https://www.phpdocx.com/img/logo_badge.png',
    )
);

$pptx->savePptx(__DIR__ . '/example_replaceVariableImage_2');
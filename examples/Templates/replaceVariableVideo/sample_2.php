<?php
// replace video variables (placeholders) using ${} to wrap placeholders. The placeholders have been added to the alt text

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_symbols.pptx');
$pptx->setTemplateSymbol('${', '}');

// replace videos in slides target
$pptx->replaceVariableVideo(
    array(
        'VIDEO_1' => __DIR__ . '/../../files/video.mp4',
    )
);

$pptx->savePptx(__DIR__ . '/example_replaceVariableVideo_2');
<?php
// replace video variables (placeholders) keeping the placeholder image. The placeholders have been added to the alt text

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// replace videos in slides target
$pptx->replaceVariableVideo(
    array(
        'VIDEO_1' => __DIR__ . '/../../files/video.mkv',
    ),
    array(
        'image' => array(
            'usePlaceholderImage' => true,
        )
    )
);

$pptx->savePptx(__DIR__ . '/example_replaceVariableVideo_4');
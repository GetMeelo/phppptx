<?php
// replace video variables (placeholders) adding a custom image. The placeholders have been added to the alt text

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// replace videos in slides target
$pptx->replaceVariableVideo(
    array(
        'VIDEO_1' => __DIR__ . '/../../files/video.mp4',
    ),
    array(
        'image' => array(
            'image' => __DIR__ . '/../../files/image.png',
        )
    )
);

$pptx->savePptx(__DIR__ . '/example_replaceVariableVideo_3');
<?php
// replace audio variables (placeholders) keeping the placeholder image. The placeholders have been added to the alt text

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// replace audios in slides target
$pptx->replaceVariableAudio(
    array(
        'AUDIO_1' => __DIR__ . '/../../files/audio.mp3',
    ),
    array(
        'image' => array(
            'usePlaceholderImage' => true,
        )
    )
);

$pptx->savePptx(__DIR__ . '/example_replaceVariableAudio_4');
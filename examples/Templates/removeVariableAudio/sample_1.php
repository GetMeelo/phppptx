<?php
// remove audio variables (placeholders). The placeholders have been added to the alt text

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// remove audios in slides target
$pptx->removeVariableAudio(array('AUDIO_1'));

$pptx->savePptx(__DIR__ . '/example_removeVariableAudio_1');
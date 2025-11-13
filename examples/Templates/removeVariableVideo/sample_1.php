<?php
// remove video variables (placeholders). The placeholders have been added to the alt text

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// remove videos in slides target
$pptx->removeVariableVideo(array('VIDEO_1'));

$pptx->savePptx(__DIR__ . '/example_removeVariableVideo_1');
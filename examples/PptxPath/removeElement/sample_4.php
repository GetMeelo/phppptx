<?php
// remove audios and videos

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// change the active slide position
$pptx->setActiveSlide(array('position' => 4));

$referenceNode = array(
    'type' => 'audio',
);
$pptx->removeElement($referenceNode);

$referenceNode = array(
    'type' => 'video',
);
$pptx->removeElement($referenceNode);

$pptx->savePptx(__DIR__ . '/example_removeElement_4');
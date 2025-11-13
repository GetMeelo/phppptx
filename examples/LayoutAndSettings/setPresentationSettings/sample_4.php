<?php
// set the presentation as read only

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();
$pptx->setPresentationSettings(array('readOnly' => true));

$pptx->savePptx(__DIR__ . '/example_setPresentationSettings_4');
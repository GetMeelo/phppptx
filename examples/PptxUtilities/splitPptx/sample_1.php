<?php
// split an existing PPTX

require_once __DIR__ . '/../../../classes/PptxUtilities.php';

$pptx = new PptxUtilities();
$pptx->splitPptx(__DIR__ . '/../../files/data_powerpoint.pptx', __DIR__ . '/splitPptx_.pptx');
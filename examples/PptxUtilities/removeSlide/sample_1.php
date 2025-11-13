<?php
// remove slides 1 and 3 from an existing PPTX

require_once __DIR__ . '/../../../classes/PptxUtilities.php';

$pptx = new PptxUtilities();
$pptx->removeSlide(__DIR__ . '/../../files/data_powerpoint.pptx', __DIR__ . '/example_removeSlide_1.pptx', array('slideNumber' => array(1, 3)));
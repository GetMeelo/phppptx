<?php
// replace string values in the first slide from an existing PPTX

require_once __DIR__ . '/../../../classes/PptxUtilities.php';

$pptx = new PptxUtilities();

$data = array('Welcome to PowerPoint' => 'Welcome to Phppptx');

$pptx->searchAndReplace(__DIR__ . '/../../files/data_powerpoint.pptx', __DIR__ . '/example_searchAndReplace_1.pptx', $data, array('slideNumber' => 1));
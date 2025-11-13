<?php
// replace string values in all slides from an existing PPTX

require_once __DIR__ . '/../../../classes/PptxUtilities.php';

$pptx = new PptxUtilities();

$data = array(
    'PowerPoint' => 'Phppptx',
    'beautiful presentations' => 'awesome presentations',
);

$pptx->searchAndReplace(__DIR__ . '/../../files/data_powerpoint.pptx', __DIR__ . '/example_searchAndReplace_2.pptx', $data);
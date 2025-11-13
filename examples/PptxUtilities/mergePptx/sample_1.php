<?php
// merge PPTX

require_once __DIR__ . '/../../../classes/MergePptx.php';

$pptx = new MergePptx();
$pptx->merge(array(__DIR__ . '/../../files/sample.pptx', __DIR__ . '/../../files/data_powerpoint.pptx', __DIR__ . '/../../files/charts.pptx', __DIR__ . '/../../files/sample_template_multi.pptx', __DIR__ . '/../../files/sample_template.pptx', __DIR__ . '/../../files/sample_comments.pptx'), __DIR__ . '/example_mergePptx_1.pptx');
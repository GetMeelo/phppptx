<?php
// transform a PPTX to HTML

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$transformHtmlPlugin = new TransformNativeHtmlDefaultPlugin();

$transform = new TransformNativeHtml(__DIR__ . '/../../files/sample_template.pptx');
$html = $transform->transform($transformHtmlPlugin);

echo $html;
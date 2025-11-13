<?php
// transform a PPTX to PPT and ODP using the conversion plugin based on LibreOffice

require_once __DIR__ . '/../../../../classes/CreatePptx.php';

$pptx = new CreatePptx();
$pptx->transform(__DIR__ . '/../../../files/sample.pptx', __DIR__ . '/transform_libreoffice_2.ppt', 'libreoffice');
$pptx->transform(__DIR__ . '/../../../files/sample.pptx', __DIR__ . '/transform_libreoffice_2.odp', 'libreoffice');
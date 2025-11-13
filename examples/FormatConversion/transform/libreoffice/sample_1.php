<?php
// transform a PPTX to PDF using the conversion plugin based on LibreOffice

require_once __DIR__ . '/../../../../classes/CreatePptx.php';

$pptx = new CreatePptx();
$pptx->transform(__DIR__ . '/../../../files/sample.pptx', __DIR__ . '/transform_libreoffice_1.pdf', 'libreoffice');
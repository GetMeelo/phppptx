<?php
// transform a PPTX to PDF using the conversion plugin based on MS PowerPoint

require_once __DIR__ . '/../../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

// global paths must be used
$pptx->transform('E:\\phppptx\\examples\\files\\sample.pptx', 'E:\\phppptx\\examples\\FormatConversion\\transform\\mspowerpoint\\transform_mspowerpoint_1.pdf', 'mspowerpoint');
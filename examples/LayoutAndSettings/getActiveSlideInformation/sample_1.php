<?php
// get the internal active slide information in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$activeSlideInformation = $pptx->getActiveSlideInformation();
var_dump($activeSlideInformation);
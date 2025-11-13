<?php
// add background images in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');

$pptx->setActiveSlide(array('position' => 0));
$pptx->addBackgroundImage(__DIR__ . '/../../files/image.png', array('transparency' => 60));

$pptx->setActiveSlide(array('position' => 2));
$pptx->addBackgroundImage(__DIR__ . '/../../files/image.png', array('tilePictureAsTexture' => true, 'transparency' => 80));

$pptx->savePptx(__DIR__ . '/example_addBackgroundImage_2');
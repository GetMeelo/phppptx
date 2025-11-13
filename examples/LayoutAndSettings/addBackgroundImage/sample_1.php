<?php
// add background images in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$pptx->addBackgroundImage(__DIR__ . '/../../files/image.png');

$pptx->addSlide(array('active' => true));
$pptx->addBackgroundImage(__DIR__ . '/../../files/image.png', array('transparency' => 60));

$pptx->addSlide(array('active' => true));
$pptx->addBackgroundImage(__DIR__ . '/../../files/image.png', array('tilePictureAsTexture' => true, 'transparency' => 80));

$pptx->savePptx(__DIR__ . '/example_addBackgroundImage_1');
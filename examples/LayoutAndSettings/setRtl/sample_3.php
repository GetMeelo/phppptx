<?php
// add HTML rtl contents in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();
$pptx->setPresentationSettings(array('rtl' => true));
$pptx->setRtl();

$pptx->addSlide(array('layout' => 'Title and Content', 'active' => true));

$html = '<p style="font-size: 48px; dir: rtl; text-align: right;"><b>Lorem ipsum</b></p>';
$pptx->addHtml($html, array('placeholder' => array('name' => 'Title 1')));

$pptx->savePptx(__DIR__ . '/example_setRtl_3');
<?php
// add slideNumber footer slide in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$content = array(
    'text' => 'My title',
);
// the getActiveSlideInformation method returns information about placeholders in the current active slide to add text contents
$pptx->addText($content, array('placeholder' => array('name' => 'Title 1')));

$pptx->addFooterSlide('slideNumber');

$pptx->savePptx(__DIR__ . '/example_addFooterslide_2');
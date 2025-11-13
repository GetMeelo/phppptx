<?php
// add text content footer slide in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$content = array(
    'text' => 'My title',
);
// the getActiveSlideInformation method returns information about placeholders in the current active slide to add text contents
$pptx->addText($content, array('placeholder' => array('name' => 'Title 1')));

$textContent = array(
    'text' => 'Footer content',
    'bold' => true,
);
$textFragment = new PptxFragment();
$textFragment->addText($textContent, array());

$pptx->addSlide(array('active' => true));

$pptx->addFooterSlide('textContents', array('textContents' => $textFragment));

$pptx->savePptx(__DIR__ . '/example_addFooterslide_3');
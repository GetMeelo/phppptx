<?php
// add footers to all slides in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$content = array(
    'text' => 'My title',
);
$pptx->addText($content, array('placeholder' => array('name' => 'Title 1')));

// create new slides
$pptx->addSlide();
$pptx->addSlide();
$pptx->addSlide();

// add fixed date to all slides
$pptx->addFooterSlide('dateAndTime', array('applyToAll' => true));
// add slide number to all slides
$pptx->addFooterSlide('slideNumber', array('applyToAll' => true));

// add text content to all slides
$textContent = array(
    'text' => 'Footer content',
    'bold' => true,
);
$textFragment = new PptxFragment();
$textFragment->addText($textContent, array());

$pptx->addFooterSlide('textContents', array('applyToAll' => true, 'textContents' => $textFragment));

$pptx->savePptx(__DIR__ . '/example_addFooterslide_4');
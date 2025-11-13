<?php
// add dateAndTime and slideNumber footer slide applying styles in a PPTX created from scratch

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

$dateAndTimeStyles = array(
    'bold' => true,
    'color' => '628A54',
    'font' => 'Times New Roman',
    'italic' => true,
);

// add date
$pptx->addFooterSlide('dateAndTime', array('applyToAll' => true, 'contentStyles' => $dateAndTimeStyles));

$slideNumberStyles = array(
    'bold' => true,
    'fontSize' => 24,
    'align' => 'center',
);

// add slide number
$pptx->addFooterSlide('slideNumber', array('applyToAll' => true, 'contentStyles' => $slideNumberStyles));

$pptx->savePptx(__DIR__ . '/example_addFooterslide_6');
<?php
// add dateAndTime footer slide in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$content = array(
    'text' => 'My title',
);
$pptx->addText($content, array('placeholder' => array('name' => 'Title 1')));

// add automatic date
$pptx->addFooterSlide('dateAndTime');

// create new slide and set it as the internal active slide
$pptx->addSlide(array('active' => true));

// add automatic date setting a custom type
$pptx->addFooterSlide('dateAndTime', array('dateAndTimeType' => 'datetime8'));

// create new slide and set it as the internal active slide
$pptx->addSlide(array('active' => true));
// add fixed date
$pptx->addFooterSlide('dateAndTime', array('dateAndTime' => '2023/12/02'));

$pptx->savePptx(__DIR__ . '/example_addFooterslide_1');
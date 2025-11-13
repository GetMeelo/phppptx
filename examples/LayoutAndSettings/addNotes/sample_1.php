<?php
// add notes in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

// add notes in the first slide
$pptx->addNotes('A new note.');
$pptx->addNotes('A new note 2.');

// add two slides
$pptx->addSlide();
$pptx->addSlide();

// add notes with styles in the third slide
$pptx->setActiveSlide(array('position' => 2));
$content = array(
    array(
        'text' => 'A note',
        'italic' => true,
    ),
    array(
        'text' => ' with ',
        'bold' => true,
        'font' => 'Times New Roman',
    ),
    array(
        'text' => 'styles.',
        'italic' => true,
        'underline' => true,
    ),
);
$pptx->addNotes($content, array('align' => 'center'));

$pptx->savePptx(__DIR__ . '/example_addNotes_1');
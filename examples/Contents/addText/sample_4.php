<?php
// add text contents in layout placeholders in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample.pptx');

$pptx->setActiveSlide(array('position' => 0));

$content = array(
    'text' => 'PPTX files.',
    'fontSize' => 18,
    'italic' => true,
);
$paragraphStyles = array(
    'align' => 'right',
    'lineSpacing' => 2,
);
$pptx->addText($content, array('placeholder' => array('name' => 'Subtitle 2')), $paragraphStyles);

$pptx->setActiveSlide(array('position' => 2));

$content = array(
    'text' => 'More products',
);
$pptx->addText($content, array('placeholder' => array('descr' => 'product_title')), array(), array('insertMode' => 'replace'));

$pptx->savePptx(__DIR__ . '/example_addText_4');
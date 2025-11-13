<?php
// remove a shape in the internal active slide in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$pptx->removeShapeSlide(array('placeholder' => array('name' => 'Subtitle 2')));

$pptx->savePptx(__DIR__ . '/example_removeShapeSlide_1');
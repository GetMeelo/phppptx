<?php
// insert math equations from OMML

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$pptx->addMathEquation('<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>∪±∞=~×</m:t></m:r></m:oMath></m:oMathPara>', 'omml', array('placeholder' => array('name' => 'Subtitle 2')));

$pptx->savePptx(__DIR__ . '/example_addMathEquation_1');
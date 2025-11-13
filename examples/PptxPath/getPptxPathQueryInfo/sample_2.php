<?php
// return the OOXML information of image elements

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// change the active slide position
$pptx->setActiveSlide(array('position' => 2));

$referenceNode = array(
    'type' => 'image',
);
$queryInfo = $pptx->getPptxPathQueryInfo($referenceNode);
var_dump($queryInfo);

if ($queryInfo['elements']->length > 0) {
    foreach ($queryInfo['elements'] as $element) {
        var_dump($element->ownerDocument->saveXML($element));
    }
}
<?php
// return the OOXML information of paragraph elements from all slides

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

// Indexer can be used to get the number of slides
$indexer = new Indexer(__DIR__ . '/../../files/sample_template.pptx');
$infoIndexer = $indexer->getOutput();

$referenceNode = array(
    'type' => 'paragraph',
);
// iterate all slides to get paragraphs from each slide
for ($iSlide = 0; $iSlide < count($infoIndexer['slides']); $iSlide++) {
    $pptx->setActiveSlide(array('position' => $iSlide));
    $queryInfo = $pptx->getPptxPathQueryInfo($referenceNode);
    var_dump($queryInfo);

    if ($queryInfo['elements']->length > 0) {
        foreach ($queryInfo['elements'] as $element) {
            var_dump($element->ownerDocument->saveXML($element));
        }
    }
}
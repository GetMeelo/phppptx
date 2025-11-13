<?php
// add comments in an existing PPTX

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_comments.pptx');

$pptx->addCommentAuthor('phppptx');

$pptx->setActiveSlide(array('position' => 1));

$pptx->addComment('New comment', 'phppptx');

$pptx->savePptx(__DIR__ . '/example_addComment_3');
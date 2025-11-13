<?php
// replace table variables (placeholders)

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

$contentTextA = array(
    array(
        'text' => 'Title ',
        'bold' => true,
        'font' => 'Arial',
        'fontSize' => 30,
    ), array(
        'text' => 'A',
        'color' => '406094',
        'fontSize' => 40,
    )
);
$textFragmentA = new PptxFragment();
$textFragmentA->addText($contentTextA, array());

$contentTextB = array(
    array(
        'text' => 'Title ',
        'bold' => true,
        'font' => 'Arial',
        'fontSize' => 30,
    ), array(
        'text' => 'B',
        'color' => '406094',
        'fontSize' => 40,
    )
);
$textFragmentB = new PptxFragment();
$textFragmentB->addText($contentTextB, array());

$contentLink = array(
    'text' => 'My link',
    'hyperlink' => 'https://www.phppptx.com'
);
$linkFragment = new PptxFragment();
$linkFragment->addText($contentLink, array());

$htmlFragment = new PptxFragment();
$html = '<p style="font-size: 24px;"><b><em>HTML</em> content</b>: Cell 3.3</p>';
$htmlFragment->addHtml($html, array());
$htmlFragment->addText($contentLink, array());

$items = array(
    array(
        'PRODUCT' => 'Product A',
        'TITLE' => $textFragmentA,
        'DESCRIPTION' => $htmlFragment,
    ),
    array(
        'PRODUCT' => 'Product B',
        'TITLE' => 'Title B',
        'DESCRIPTION' => $linkFragment,
    ),
);

$pptx->replaceVariableTable($items);

$pptx->savePptx(__DIR__ . '/example_replaceVariableTable_3');
<?php
// replace list variables (placeholders) with new PPTXFragments

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template.pptx');

$contentA = array(
    'text' => 'List item 1',
    'bold' => true,
    'underline' => 'single',
);
$listFragmentA = new PptxFragment();
$listFragmentA->addText($contentA, array());

$contentB = array(
    array(
        'text' => 'Sublist',
        'highlight' => '628A54',
    ),
    array(
        'text' => ' item 1.1',
        'bold' => true,
        'italic' => true,
        'font' => 'Times New Roman',
    ),
);
$listFragmentB = new PptxFragment();
$listFragmentB->addText($contentB, array());

$contentC = array(
    array(
        'text' => 'List item ',
        'bold' => true,
        'font' => 'Arial',
        'fontSize' => 30,
    ), array(
        'text' => '3',
        'color' => '406094',
        'fontSize' => 40,
    )
);
$listFragmentC = new PptxFragment();
$listFragmentC->addText($contentC, array());

$contentLink = array(
    'text' => 'My link',
    'hyperlink' => 'https://www.phppptx.com'
);
$linkFragment = new PptxFragment();
$linkFragment->addText($contentLink, array());

$htmlFragment = new PptxFragment();
$html = '<p style="font-size: 42px;"><b><em>HTML</em> content</b>: Cell 3.3</p>';
$htmlFragment->addHtml($html, array());

$items = array(
    $listFragmentA,
    array(
        $listFragmentB,
    ),
    $listFragmentC,
    $linkFragment,
    $htmlFragment,
);

$pptx->replaceVariableList('VAR_LIST', $items);

$pptx->savePptx(__DIR__ . '/example_replaceVariableList_3');
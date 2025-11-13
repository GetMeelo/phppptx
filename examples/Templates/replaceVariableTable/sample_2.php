<?php
// replace table variables (placeholders) using ${} to wrap placeholders

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptxFromTemplate(__DIR__ . '/../../files/sample_template_symbols.pptx');
$pptx->setTemplateSymbol('${', '}');

$items = array(
    array(
        'ITEM' => 'Product A',
        'REFERENCE' => '107AW3',
        'DETAILS' => 'Details A',
    ),
    array(
        'ITEM' => 'Product B',
        'REFERENCE' => '204RS67O',
        'DETAILS' => 'Details B',
    ),
    array(
        'ITEM' => 'Product C',
        'REFERENCE' => '25GTR56',
        'DETAILS' => 'Details C',
    )
);

$pptx->replaceVariableTable($items);

$items = array(
    array(
        'PRODUCT' => 'Product A',
        'TITLE' => 'Title A',
        'DESCRIPTION' => 'Desc A',
    ),
    array(
        'PRODUCT' => 'Product B',
        'TITLE' => 'Title B',
        'DESCRIPTION' => 'Desc B',
    ),
);

$pptx->replaceVariableTable($items);

$pptx->savePptx(__DIR__ . '/example_replaceVariableTable_2');
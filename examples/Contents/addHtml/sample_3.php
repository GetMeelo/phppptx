<?php
// add HTML list contents in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$position = array(
    'coordinateX' => 4500000,
    'coordinateY' => 4000000,
    'sizeX' => 2000000,
    'sizeY' => 2000000,
);
$html = '<style>
    p {
        font-family: Arial;
        font-style: italic;
        text-decoration: underline;
        color: #2CA87A;
    }
    ul {
        font-weight: bold;
    }
</style>
<p>Unordered list:</p>
<ul>
    <li>Item 1</li>
    <li>Item 2</li>
    <li>Item 3</li>
</ul>';
$pptx->addHtml($html, array('new' => $position));

$position = array(
    'coordinateX' => 500000,
    'coordinateY' => 400000,
    'sizeX' => 3000000,
    'sizeY' => 2450000,
    'textBoxStyles' => array(
        'border' => array(
            'color' => 'FF0000',
            'dash' => 'dash',
            'width' => 48575,
        ),
    ),
);
$html = '<style>
    p {
        font-family: Arial;
        font-style: italic;
        text-decoration: underline;
        color: #2CA87A;
    }
    ol.rls {
        list-style-type: checkmarkBullet;
    }
    ul.sls {
        list-style-type: square;
    }
    ul.cls {
        list-style-type: circle;
    }
</style>
<p>Ordered list:</p>
<ol>
    <li>Item 1</li>
    <li>Item 2</li>
    <ol class="rls">
        <li>Item 2.1</li>
        <li>Item 2.2</li>
        <li>Item 2.3</li>
    </ol>
    <li>Item 3</li>
    <ul class="sls">
        <li>Item 3.1</li>
        <li>Item 3.2</li>
        <li>Item 3.3
        <ul class="cls">
            <li>Item 3.3.1</li>
            <li>Item 3.3.2</li>
            <li>Item 3.3.3</li>
        </ul>
        </li>
    </ul>
    <li>Item 4</li>
</ol>';
$pptx->addHtml($html, array('new' => $position));

$pptx->savePptx(__DIR__ . '/example_addHtml_3');
<?php
// add HTML contents in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$html = '<p style="font-size: 48px;"><b>Lorem ipsum</b></p>';
$pptx->addHtml($html, array('placeholder' => array('name' => 'Title 1')));

$html = '<style>
    p {
        font-family: Arial;
        font-style: italic;
        text-decoration: underline;
        color: #2CA87A;
    }
</style>
<p>Lorem ipsum with <br>Arial font</p>';
$pptx->addHtml($html, array('placeholder' => array('name' => 'Subtitle 2')));

$position = array(
    'coordinateX' => 770000,
    'coordinateY' => 507000,
    'sizeX' => 7000000,
    'sizeY' => 450000,
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
        background-color: #AEBF7A;
    }
    .c1 {
        text-decoration: underline;
        font-size: 16pt;
    }
</style>
<p>
    <b>phppptx</b> can transform <span class="c1">HTML to PPTX</span>.
    <a href="https://www.phppptx.com">Link to phppptx</a>
</p>
';
$pptx->addHtml($html, array('new' => $position));

$pptx->savePptx(__DIR__ . '/example_addHtml_1');
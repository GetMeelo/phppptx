<?php
// add HTML contents in a PPTX created from scratch

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$html = '<p style="font-size: 48px;"><b>Lorem ipsum</b></p>';
$pptx->addHtml($html, array('placeholder' => array('name' => 'Title 1')));

$html = '<style>
:root {
    --blue: #1990fa;
    font-weight: bold;
}
.varcolor {
    color: var(--blue);
}
.color {
    color: green;
}
</style>
<p class="varcolor">Text with blue color using a CSS variable.</p>
<p class="color">Text with a custom color.</p>';
$pptx->addHtml($html, array('placeholder' => array('name' => 'Subtitle 2')), array('parseCSSVars' => true));

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
:root {
    --text-color: green;
    --myborders: dashed;
    --text-size: 4px;
    font-weight: bold;
    --color-r: red;
    --color-g: green;
    --sz-t1: 32px;
    --font-fc: Arial;
}

p {
    border-bottom: var(--text-size) var(--myborders) var(--text-color);
    color: var(--color-r);
    font-size: var(--sz-t1, 32px);
}

span.c1 {
    color: var(--color-g);
    font-size: var(--sz-t2, 16px);
    font-family: var(--font-fc);
}

</style>
<p>Complex <span class="c1">CSS variables</span> can be used.</p>
';
$pptx->addHtml($html, array('new' => $position), array('parseCSSVars' => true));

$pptx->savePptx(__DIR__ . '/example_addHtml_2');
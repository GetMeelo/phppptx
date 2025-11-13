<?php
// add an SVG content

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

$position = array(
    'coordinateX' => 1400000,
    'coordinateY' => 4500000,
);

$svg = '
<svg height="100" width="100">
  <circle cx="50" cy="50" r="40" stroke="black" stroke-width="3" fill="red" />
</svg>';

$pptx->addSvg($svg, $position);

$pptx->savePptx(__DIR__ . '/example_addSvg_2');
<?php
// insert a math equation from MathML applying styles

require_once __DIR__ . '/../../../classes/CreatePptx.php';

$pptx = new CreatePptx();

// text box position and styles
$position = array(
    'coordinateX' => 4500000,
    'coordinateY' => 5500000,
    'sizeX' => 7000000,
    'sizeY' => 450000,
);

$mathML = '<math xmlns="http://www.w3.org/1998/Math/MathML">
	<mrow>
		<mi>A</mi>
		<mo>=</mo>
		<mfenced open="[" close="]">
			<mtable>
				<mtr>
					<mtd>
						<mi>x</mi>
					</mtd>
					<mtd>
						<mn>2</mn>
					</mtd>
				</mtr>
				<mtr>
					<mtd>
						<mn>3</mn>
					</mtd>
					<mtd>
						<mi>w</mi>
					</mtd>
				</mtr>
			</mtable>
		</mfenced>
	</mrow>
</math>';

$pptx->addMathEquation($mathML, 'mathml', array('new' => $position), array('align' => 'right', 'color' => 'FF0000', 'fontSize' => 32));

$pptx->savePptx(__DIR__ . '/example_addMathEquation_3');
<?php

/**
 * Use an existing PPTX as the document template
 *
 * @category   Phppptx
 * @package    create
 * @copyright  Copyright (c) Narcea Labs SL
 *             (https://www.narcealabs.com)
 * @license    phppptx LICENSE
 * @link       https://www.phppptx.com
 */
class CreatePptxFromTemplate extends CreatePptx
{
    /**
     *
     * @access public
     * @var bool
     */
    public $preprocessed = false;

    /**
     *
     * @access public
     * @var string
     */
    public $templateSymbolStart = '$';

    /**
     *
     * @access public
     * @var string
     */
    public $templateSymbolEnd = '$';

    /**
     *
     * @access public
     * @static
     * @var string
     */
    public static $regExprVariableSymbols = '([^ ]*)';

    /**
     *
     * @access public
     * @static
     * @var string
     */
    public $templateBlockSymbol = 'BLOCK_';

    /**
     * Construct
     *
     * @access public
     * @param mixed $pptxTemplatePath path to the template to use or PptxStructure
     * @param array $options
     *      'preprocessed' (bool) if true the variables will not be 'repaired'. Default as false
     * @throws Exception empty or not valid template
     */
    public function __construct($pptxTemplatePath, $options = array())
    {
        if (empty($pptxTemplatePath)) {
            PhppptxLogger::logger('The template path can not be empty', 'fatal');
        }

        // default options
        $this->preprocessed = false;

        parent::__construct($options, $pptxTemplatePath);

        if (isset($options['preprocessed']) && $options['preprocessed']) {
            $this->preprocessed = true;
        }
    }

    /**
     * Getter. Return preprocessed
     *
     * @access public
     * @return bool
     */
    public function getPreprocessed()
    {
        return $this->preprocessed;
    }

    /**
     * Getter. Return template symbol
     *
     * @access public
     * @return string|array
     */
    public function getTemplateSymbol()
    {
        if ($this->templateSymbolStart == $this->templateSymbolEnd && strlen($this->templateSymbolStart) == 1) {
            return $this->templateSymbolStart;
        } else {
            return array($this->templateSymbolStart, $this->templateSymbolEnd);
        }
    }

    /**
     * Setter. Set preprocessed
     *
     * @access public
     * @param bool $preprocessed
     */
    public function setPreprocessed($preprocessed)
    {
        $this->preprocessed = $preprocessed;
    }

    /**
     * Setter. Set template symbol
     *
     * @access public
     * @param string $templateSymbolStart
     * @param string $templateSymbolEnd use $templateSymbolStart if null
     * @param array $options
     *      'preprocessed' (bool) set the template as preprocessed. Default as false
     */
    public function setTemplateSymbol($templateSymbolStart = '$', $templateSymbolEnd = null, $options = array())
    {
        if (is_null($templateSymbolEnd)) {
            $this->templateSymbolStart = $templateSymbolStart;
            $this->templateSymbolEnd = $templateSymbolStart;
        } else {
            $this->templateSymbolStart = $templateSymbolStart;
            $this->templateSymbolEnd = $templateSymbolEnd;
        }

        $this->preprocessed = false;

        if (isset($options['preprocessed']) && $options['preprocessed']) {
            $this->preprocessed = true;
        }

        PhppptxLogger::logger('Set new template symbol.', 'info');
    }

    /**
     * Returns the template variables
     *
     * @access public
     * @param string $target may be all (default), notesSlides, slides, slideLayouts, slideMasters
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     * @param array $variables
     * @return array
     */
    public function getTemplateVariables($target = 'all', $options = array(), $variables = array())
    {
        if (!$this->preprocessed) {
            $this->repairVariables();
        }

        $targetTypes = array('notesSlides', 'slides', 'slideLayouts', 'slideMasters');

        if ($target == 'all') {
            foreach ($targetTypes as $targets) {
                $variables = $this->getTemplateVariables($targets, $options, $variables);
            }
        } else {
            $targetContents = $this->zipPptx->getContentByType($target);
            if ($target == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
                $targetContents = array($targetContents[$this->activeSlide['position']]);
            }

            foreach ($targetContents as $targetContent) {
                $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);
                $targetXPath = new DOMXPath($targetDOM);
                $targetXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

                // strings
                // iterate a:t tags
                $nodesT = $targetDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't');

                foreach ($nodesT as $nodeT) {
                    $newVariables = $this->extractVariables($nodeT->nodeValue);
                    foreach ($newVariables as $newVariable) {
                        if (!isset($variables[$target]) || !in_array($newVariable, $variables[$target])) {
                            $variables[$target][] = $newVariable;
                        }
                    }
                }

                // pics
                $nodesPic = $targetXPath->query('//p:pic//p:cNvPr[@title or @descr]');
                foreach ($nodesPic as $nodePic) {
                    $newVariables = array();
                    if ($nodePic->hasAttribute('title')) {
                        $newVariables = $this->extractVariables($nodePic->getAttribute('title'));
                    }
                    if ($nodePic->hasAttribute('descr')) {
                        $newVariables = $this->extractVariables($nodePic->getAttribute('descr'));
                    }
                    foreach ($newVariables as $newVariable) {
                        if (!isset($variables[$target]) || !in_array($newVariable, $variables[$target])) {
                            $variables[$target][] = $newVariable;
                        }
                    }
                }

                // free DOMDocument resources
                $targetDOM = null;
            }
        }

        PhppptxLogger::logger('Get template variables.', 'info');

        return $variables;
    }

    /**
     * Returns the template variables and their types
     *
     * @access public
     * @param string $target may be all (default), notesSlides, slides, slideLayouts, slideMasters
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     * @param array $variables
     * @return array
     * @throws Exception method not available
     */
    public function getTemplateVariablesType($target = 'all', $options = array(), $variables = array())
    {
        if (!file_exists(__DIR__ . '/PptxPathStyles.php')) {
            PhppptxLogger::logger('This method is not available for your license.', 'fatal');
        }

        // get existing variables to analyze them
        $variables = $this->getTemplateVariables($target, $options, $variables);
        $variablesTypes = array();
        $pptxpathStyles = new PptxPathStyles();

        // iterate variables to get their types
        foreach ($variables as $variablesTargetKey => $variablesTargetValue) {
            if (is_array($variablesTargetValue) && count($variablesTargetValue) > 0) {
                $variablesTypes[$variablesTargetKey] = array();
                foreach ($variablesTargetValue as $variableTargetValue) {
                    $variableTargetValueSymbols = $this->templateSymbolStart . $variableTargetValue . $this->templateSymbolEnd;

                    $targetContents = $this->zipPptx->getContentByType($variablesTargetKey);
                    if ($variablesTargetKey == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
                        $targetContents = array($targetContents[$this->activeSlide['position']]);
                    }

                    foreach ($targetContents as $targetContent) {
                        $variableType = $pptxpathStyles->analyzeVariable($variableTargetValueSymbols, $targetContent['content']);

                        if (count($variableType) > 0) {
                            $variablesTypes[$variablesTargetKey][] = array(
                                'variable' => $variableTargetValue,
                                'type' => $variableType,
                            );
                        }
                    }
                }
            }
        }

        return $variablesTypes;
    }

    /**
     * Processes the template
     *
     * @access public
     */
    public function processTemplate()
    {
        // repairVariables detects and cleans variables automatically
        $this->repairVariables();
    }

    /**
     * Removes template audio variables
     *
     * @access public
     * @param array $variables Variables to be removed
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'target' (string) slides (default), slideLayouts, slideMasters
     */
    public function removeVariableAudio($variables, $options = array())
    {
        // default values
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }

        $targetContents = $this->zipPptx->getContentByType($options['target']);
        if ($options['target'] == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
            $targetContents = array($targetContents[$this->activeSlide['position']]);
        }

        foreach ($targetContents as $targetContent) {
            $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);;
            $targetXPath = new DOMXPath($targetDOM);
            $targetXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
            foreach ($variables as $variable) {
                $this->removeMedia($variable, $targetDOM, $targetXPath, 'audio', $options);
            }

            // refresh contents
            $this->zipPptx->addContent($targetContent['path'], $targetDOM->saveXML());

            // free DOMDocument resources
            $targetDOM = null;
        }

        PhppptxLogger::logger('Remove template audio variable.', 'info');
    }

    /**
     * Removes template image variables
     *
     * @access public
     * @param array $variables Variables to be removed
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'target' (string) slides (default), slideLayouts, slideMasters
     */
    public function removeVariableImage($variables, $options = array())
    {
        // default values
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }

        $targetContents = $this->zipPptx->getContentByType($options['target']);
        if ($options['target'] == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
            $targetContents = array($targetContents[$this->activeSlide['position']]);
        }

        foreach ($targetContents as $targetContent) {
            $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);;
            $targetXPath = new DOMXPath($targetDOM);
            $targetXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
            foreach ($variables as $variable) {
                $this->removeMedia($variable, $targetDOM, $targetXPath, 'image', $options);
            }

            // refresh contents
            $this->zipPptx->addContent($targetContent['path'], $targetDOM->saveXML());

            // free DOMDocument resources
            $targetDOM = null;
        }

        PhppptxLogger::logger('Remove template image variable.', 'info');
    }

    /**
     * Removes template text variables
     *
     * @access public
     * @param array $variables Variables to be removed
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'target' (string) slides (default), notesSlides, slideLayouts, slideMasters
     *      'type' (string) inline (default, remove only the variable), block (remove the variable and its containing paragraph)
     */
    public function removeVariableText($variables, $options = array())
    {
        // set empty values for all variables to be used with replaceVariableText
        $variablesFilled = array();
        foreach ($variables as $variable) {
            $variablesFilled[$variable] = '';
        }
        // remove paragraph
        if (isset($options['type']) && $options['type'] == 'block') {
            $options['remove'] = true;
        }

        PhppptxLogger::logger('Remove template variable.', 'info');

        $this->replaceVariableText($variablesFilled, $options);
    }

    /**
     * Removes template video variables
     *
     * @access public
     * @param array $variables Variables to be removed
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'target' (string) slides (default), slideLayouts, slideMasters
     */
    public function removeVariableVideo($variables, $options = array())
    {
        // default values
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }

        $targetContents = $this->zipPptx->getContentByType($options['target']);
        if ($options['target'] == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
            $targetContents = array($targetContents[$this->activeSlide['position']]);
        }

        foreach ($targetContents as $targetContent) {
            $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);;
            $targetXPath = new DOMXPath($targetDOM);
            $targetXPath->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
            foreach ($variables as $variable) {
                $this->removeMedia($variable, $targetDOM, $targetXPath, 'video', $options);
            }

            // refresh contents
            $this->zipPptx->addContent($targetContent['path'], $targetDOM->saveXML());

            // free DOMDocument resources
            $targetDOM = null;
        }

        PhppptxLogger::logger('Remove template video variable.', 'info');
    }

    /**
     * Replaces audio placeholders by an external audio
     *
     * @access public
     * @param array $variables variable names and audio paths
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'image' (array)
     *          'image' image to be used as preview. Set a default one if not set.
     *          'mime' (string) forces a mime (image/jpg, image/jpeg, image/png, image/gif, image/bmp, image/webp)
     *          'usePlaceholderImage' (bool) if true, do not change the placeholder image. Default as false
     *      'mime' (string) forces a mime (audio/mpeg, audio/x-wav, audio/x-ms-wma, audio/unknown)
     *      'target' (string) slides (default), slideLayouts, slideMasters
     * @throws Exception audio doesn't exist
     * @throws Exception audio format is not supported
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     */
    public function replaceVariableAudio($variables, $options = array())
    {
        // default values
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }

        $targetContents = $this->zipPptx->getContentByType($options['target']);
        if ($options['target'] == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
            $targetContents = array($targetContents[$this->activeSlide['position']]);
        }

        foreach ($targetContents as $targetContent) {
            $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);

            // add the new relationship
            $targetRelsPath = str_replace('ppt/'.$options['target'].'/', 'ppt/'.$options['target'].'/_rels/', $targetContent['path']) . '.rels';
            $targetRelsDOM = $this->zipPptx->getContent($targetRelsPath, 'DOMDocument');

            foreach ($variables as $variableKey => $variableSrc) {
                $newRels = $this->replaceMedia($variableKey, $variableSrc, $targetDOM, 'audio', $options);

                if (count($newRels) > 0) {
                    foreach ($newRels as $newRel) {
                        // add the new relationship
                        $this->generateRelationship($targetRelsDOM, $newRel);
                    }
                }
            }

            // refresh contents
            $this->zipPptx->addContent($targetContent['path'], $targetDOM->saveXML());
            if ($targetRelsDOM) {
                $this->zipPptx->addContent($targetRelsPath, $targetRelsDOM->saveXML());
            }

            // free DOMDocument resources
            $targetDOM = null;
            $targetRelsDOM = null;
        }

        PhppptxLogger::logger('Replace variable audio.', 'info');
    }

    /**
     * Replaces image placeholders by an external image
     *
     * @access public
     * @param array $variables variable names and image paths
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'descr' (string) set a descr value
     *      'mime' (string) forces a mime (image/jpg, image/jpeg, image/png, image/gif, image/bmp, image/webp)
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     *      'target' (string) slides (default), slideLayouts, slideMasters
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     * @throws Exception mime option is not set and getimagesizefromstring is not available
     */
    public function replaceVariableImage($variables, $options = array())
    {
        // default values
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }

        $targetContents = $this->zipPptx->getContentByType($options['target']);
        if ($options['target'] == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
            $targetContents = array($targetContents[$this->activeSlide['position']]);
        }

        foreach ($targetContents as $targetContent) {
            $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);

            // get the relationship content
            $targetRelsPath = str_replace('ppt/'.$options['target'].'/', 'ppt/'.$options['target'].'/_rels/', $targetContent['path']) . '.rels';
            $targetRelsDOM = $this->zipPptx->getContent($targetRelsPath, 'DOMDocument');

            foreach ($variables as $variableKey => $variableSrc) {
                $newRels = $this->replaceImage($variableKey, $variableSrc, $targetDOM, $options);

                if (!empty($newRels)) {
                    // add the new relationship
                    $this->generateRelationship($targetRelsDOM, $newRels);
                }
            }

            // refresh contents
            $this->zipPptx->addContent($targetContent['path'], $targetDOM->saveXML());
            if ($targetRelsDOM) {
                $this->zipPptx->addContent($targetRelsPath, $targetRelsDOM->saveXML());
            }

            // free DOMDocument resources
            $targetDOM = null;
            $targetRelsDOM = null;
        }

        PhppptxLogger::logger('Replace variable image.', 'info');
    }

    /**
     * Replaces a single variable within a list by a list of items
     *
     * @access public
     * @param string $variable
     * @param array $listValues
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'target' (string) slides (default), slideLayouts, slideMasters
     */
    public function replaceVariableList($variable, $listValues, $options = array())
    {
        if (!$this->preprocessed) {
            $this->repairVariables();
        }

        // default values
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }

        $targetContents = $this->zipPptx->getContentByType($options['target']);
        if ($options['target'] == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
            $targetContents = array($targetContents[$this->activeSlide['position']]);
        }

        foreach ($targetContents as $targetContent) {
            $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);
            $targetXPath = new DOMXPath($targetDOM);
            $targetXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

            $searchContent = $this->templateSymbolStart . $variable . $this->templateSymbolEnd;
            $foundNodes = $targetXPath->query('//a:p[a:r/a:t[text()[contains(., "' . $searchContent . '")]]]');
            if ($foundNodes->length > 0) {
                foreach ($foundNodes as $domNode) {
                    $numPprNodes = $domNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'pPr');
                    if ($numPprNodes->length > 0 && $numPprNodes->item(0)->hasAttribute('lvl')) {
                        $this->replaceListValues($searchContent, $domNode, $listValues, (int)$numPprNodes->item(0)->hasAttribute('lvl'), $options, $targetContent);
                    } else {
                        $this->replaceListValues($searchContent, $domNode, $listValues, 0, $options, $targetContent);
                    }
                    $domNode->parentNode->removeChild($domNode);
                }
                $this->zipPptx->addContent($targetContent['path'], $targetDOM->saveXML());
            }

            // free DOMDocument resources
            $targetDOM = null;
        }

        PhppptxLogger::logger('Replace variable list.', 'info');
    }

    /**
     * Replaces an array of variables with HTML
     *
     * @access public
     * @param array $variables variable names and HTML contents
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'forceNotTidy' (bool) if true, avoid using Tidy. Only recommended if Tidy can't be installed. Default as false
     *      'parseCSSVars' (bool) parse CSS variables. Default as false
     *      'target' (string) slides (default), notesSlides, slideLayouts, slideMasters
     *      'type' (string) inline (replace only the variable), block (default, remove the variable and its containing paragraph)
     * @throws Exception not valid PptxFragment, PHP Tidy is not available and forceNotTidy is false
     */
    public function replaceVariableHtml($variables, $options = array())
    {
        // default values
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }

        foreach ($variables as $variableKey => $variableValue) {
            $htmlFragment = new PptxFragment();
            $htmlFragment->addHtml($variableValue, $options);

            $this->replaceVariablePptxFragment(array($variableKey => $htmlFragment), $options);
        }

        PhppptxLogger::logger('Replace variable HTML.', 'info');
    }

    /**
     * Replaces an array of variables with PptxFragments
     *
     * @access public
     * @param array $variables variable names and PptxFragments
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'target' (string) slides (default), notesSlides, slideLayouts, slideMasters
     *      'type' (string) inline (replace only the variable), block (default, replace the variable and its containing paragraph)
     * @throws Exception not valid PptxFragment
     */
    public function replaceVariablePptxFragment($variables, $options = array())
    {
        if (!$this->preprocessed) {
            $this->repairVariables();
        }

        // default values
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }
        if (!isset($options['type'])) {
            $options['type'] = 'block';
        }

        $targetContents = $this->zipPptx->getContentByType($options['target']);
        if ($options['target'] == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
            $targetContents = array($targetContents[$this->activeSlide['position']]);
        }

        foreach ($targetContents as $targetContent) {
            $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);
            $targetXPath = new DOMXPath($targetDOM);
            $targetXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
            $targetRelsPath = str_replace('ppt/'.$options['target'].'/', 'ppt/'.$options['target'].'/_rels/', $targetContent['path']) . '.rels';
            $targetRelsDOM = $this->zipPptx->getContent($targetRelsPath, 'DOMDocument');

            $nodesToBeRemoved = array();

            foreach ($variables as $variableKey => $variableValue) {
                if (!$variableValue instanceof PptxFragment) {
                    PhppptxLogger::logger('This method requires that the variable value is a PptxFragment', 'fatal');
                }

                if ($options['type'] == 'inline') {
                    // get inline text tag
                    $foundNodes = $targetXPath->query('//a:p/a:r[a:t[text()[contains(., "' . $this->templateSymbolStart . $variableKey . $this->templateSymbolEnd . '")]]]');
                } else {
                    // get block paragraph tag
                    $foundNodes = $targetXPath->query('//a:p[a:r/a:t[text()[contains(., "' . $this->templateSymbolStart . $variableKey . $this->templateSymbolEnd . '")]]]');
                }

                if ($foundNodes->length > 0) {
                    // handle external relationships such as hyperlinks
                    $externalRelationships = $variableValue->getExternalRelationships();
                    if (count($externalRelationships) > 0) {
                        $this->addExternalRelationships($externalRelationships, $targetContent['path'], $targetRelsDOM);
                    }

                    foreach ($foundNodes as $node) {
                        if ($options['type'] == 'block') {
                            // block type replacement

                            // import the new contents
                            $newNodeFragments = $variableValue->blockPptxXML();

                            foreach ($newNodeFragments as $newNodeFragment) {
                                $newContentFragment = $node->ownerDocument->createDocumentFragment();
                                $newContentFragment->appendXML($newNodeFragment->ownerDocument->saveXML($newNodeFragment));
                                $node->parentNode->appendChild($newContentFragment);
                            }

                            // remove paragraph node of the variable
                            $nodesToBeRemoved[] = $node;
                        } else if ($options['type'] == 'inline') {
                            // inline type replacement

                            // import the new contents
                            $newNodeFragments = $variableValue->inlinePptxXML();

                            $newNodesContent = '';
                            foreach ($newNodeFragments as $newNodeFragment) {
                                $newNodesContent .= $newNodeFragment->ownerDocument->saveXML($newNodeFragment);
                            }

                            // get a:rPr existing styles to be applied to elements after the replaced placeholder in the same paragraph
                            $stylesRpr = '';
                            $nodesRpr = $node->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'rPr');
                            if ($nodesRpr->length > 0) {
                                $stylesRpr = $nodesRpr->item(0)->ownerDocument->saveXML($nodesRpr->item(0));
                            }

                            $nodeContent = $node->ownerDocument->saveXML($node);

                            $nodeContent = str_replace($this->templateSymbolStart . $variableKey . $this->templateSymbolEnd, '</a:t></a:r>'.$newNodesContent.'<a:r>'.$stylesRpr.'<a:t>', $nodeContent);
                            $nodeContent = str_replace('<a:r>', '<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">', $nodeContent);

                            $newContentFragment = $node->ownerDocument->createDocumentFragment();
                            $newContentFragment->appendXML($nodeContent);
                            $node->parentNode->insertBefore($newContentFragment, $node->nextSibling);

                            // remove paragraph node of the variable
                            $nodesToBeRemoved[] = $node;
                        }
                    }
                }
            }

            foreach ($nodesToBeRemoved as $nodeToBeRemoved) {
                $nodeToBeRemoved->parentNode->removeChild($nodeToBeRemoved);
            }

            $this->zipPptx->addContent($targetContent['path'], $targetDOM->saveXML());
            if ($targetRelsDOM) {
                $this->zipPptx->addContent($targetRelsPath, $targetRelsDOM->saveXML());
            }

            // free DOMDocument resources
            $targetDOM = null;
            $targetRelsDOM = null;
        }

        PhppptxLogger::logger('Replace variable PptxFragment.', 'info');
    }

    /**
     * Replaces table variables in a 'table set of rows'
     *
     * @access public
     * @param array $variables
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'target' (string) slides (default), slideLayouts, slideMasters
     */
    public function replaceVariableTable($variables, $options = array())
    {
        if (!$this->preprocessed) {
            $this->repairVariables();
        }

        // default values
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }

        $targetContents = $this->zipPptx->getContentByType($options['target']);
        if ($options['target'] == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
            $targetContents = array($targetContents[$this->activeSlide['position']]);
        }

        foreach ($targetContents as $targetContent) {
            $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);
            $targetXPath = new DOMXPath($targetDOM);
            $targetXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

            $tableXPathQuery = '//a:tbl/a:tr[';
            $firstElement = true;
            foreach ($variables[0] as $variableKey => $variableValue) {
                if ($firstElement) {
                    $tableXPathQuery .= ' contains(.,"'.$this->templateSymbolStart.$variableKey.$this->templateSymbolEnd.'")';
                    $firstElement = false;
                } else {
                    $tableXPathQuery .= ' and contains(.,"'.$this->templateSymbolStart.$variableKey.$this->templateSymbolEnd.'")';
                }
            }
            $tableXPathQuery .= ']';

            $foundNodes = $targetXPath->query($tableXPathQuery);
            if ($foundNodes->length > 0) {
                foreach ($foundNodes as $node) {
                    $rowSpanTr = 0;
                    foreach ($variables as $variablesRow) {
                        $newNode = $node->cloneNode(true);
                        $textNodes = $newNode->getElementsBytagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't');
                        foreach ($textNodes as $text) {
                            $sxText = simplexml_import_dom($text);
                            $strNodeReplaced = (string)$sxText;
                            $nodesToBeRemoved = array();
                            foreach ($variablesRow as $variableKey => $variableValue) {
                                if ($variableValue instanceof PptxFragment) {
                                    // PptxFragment replacement

                                    if (strstr($text->nodeValue, $variableKey)) {
                                        // get a:rPr existing styles to be applied to elements after the replaced placeholder in the same paragraph
                                        $stylesRpr = '';

                                        // import the new contents
                                        $newNodeFragments = $variableValue->blockPptxXML();

                                        foreach ($newNodeFragments as $newNodeFragment) {
                                            $newContentFragment = $text->ownerDocument->createDocumentFragment();
                                            $newContentFragment->appendXML($newNodeFragment->ownerDocument->saveXML($newNodeFragment));
                                            $text->parentNode->parentNode->parentNode->appendChild($newContentFragment);
                                        }

                                        // remove existing a:r tags
                                        $nodesToBeRemoved[] = $text->parentNode->parentNode;

                                        // handle external relationships such as hyperlinks
                                        $externalRelationships = $variableValue->getExternalRelationships();
                                        if (count($externalRelationships) > 0) {
                                            $targetRelsPath = str_replace('ppt/'.$options['target'].'/', 'ppt/'.$options['target'].'/_rels/', $targetContent['path']) . '.rels';
                                            $targetRelsDOM = $this->zipPptx->getContent($targetRelsPath, 'DOMDocument');

                                            $this->addExternalRelationships($externalRelationships, $targetContent['path'], $targetRelsDOM);

                                            // refresh contents
                                            if ($targetRelsDOM) {
                                                $this->zipPptx->addContent($targetRelsPath, $targetRelsDOM->saveXML());
                                            }

                                            // free DOMDocument resources
                                            $targetRelsDOM = null;
                                        }
                                    }
                                } else {
                                    // text replacement

                                    $strNodeReplaced = str_replace($this->templateSymbolStart.$variableKey.$this->templateSymbolEnd, $variableValue, $strNodeReplaced);
                                }
                            }
                            $sxText[0] = $strNodeReplaced;

                            foreach ($nodesToBeRemoved as $nodeToBeRemoved) {
                                $nodeToBeRemoved->parentNode->removeChild($nodeToBeRemoved);
                            }
                        }

                        // check if some a:tc in the a:tr has a rowSpan attribute to add the new node in the correct node position
                        $foundNodesTcRowSpan = $targetXPath->query('./a:tc[@rowSpan]', $node);
                        if ($foundNodesTcRowSpan->length > 0) {
                            foreach ($foundNodesTcRowSpan as $foundNodeTcRowSpan) {
                                if ($foundNodeTcRowSpan->hasAttribute('rowSpan')) {
                                    if ((int)$foundNodeTcRowSpan->getAttribute('rowSpan') > $rowSpanTr) {
                                        $rowSpanTr = (int)$foundNodeTcRowSpan->getAttribute('rowSpan');
                                    }
                                }
                            }
                        }

                        $node->parentNode->insertBefore($newNode, $node);

                        if ($rowSpanTr > 1) {
                            $nextNode = $node->nextSibling;
                            for ($iRowSpanTr = 1; $iRowSpanTr < $rowSpanTr; $iRowSpanTr++) {
                                $newNodeNext = $nextNode->cloneNode(true);
                                $node->parentNode->insertBefore($newNodeNext, $node);
                                $nextNode = $node->nextSibling;
                            }
                        }
                    }

                    if ($rowSpanTr > 1) {
                        for ($iRowSpanTr = 1; $iRowSpanTr < $rowSpanTr; $iRowSpanTr++) {
                            $node->parentNode->removeChild($node->nextSibling);
                        }
                    }

                    $node->parentNode->removeChild($node);
                }

                $this->zipPptx->addContent($targetContent['path'], $targetDOM->saveXML());
            }

            // free DOMDocument resources
            $targetDOM = null;
        }

        PhppptxLogger::logger('Replace variable table.', 'info');
    }

    /**
     * Replaces an array of variables by their values
     *
     * @access public
     * @param array $variables variable names and new values
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'remove' (bool) if true, remove the paragraph. Used by removeVariableText. Default as false
     *      'target' (string) slides (default), notesSlides, slideLayouts, slideMasters
     *      'type' (string) inline (default, replace only the variable), block (replace the variable and its containing paragraph)
     */
    public function replaceVariableText($variables, $options = array())
    {
        if (!$this->preprocessed) {
            $this->repairVariables();
        }

        // default values
        if (!isset($options['remove'])) {
            $options['remove'] = false;
        }
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }
        if (!isset($options['type'])) {
            $options['type'] = 'inline';
        }

        $targetContents = $this->zipPptx->getContentByType($options['target']);
        if ($options['target'] == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
            $targetContents = array($targetContents[$this->activeSlide['position']]);
        }

        foreach ($targetContents as $targetContent) {
            $newContent = $targetContent['content'];

            if ($options['type'] == 'block') {
                // block replacement
                $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);
                $targetXPath = new DOMXPath($targetDOM);
                $targetXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
                foreach ($variables as $variableKey => $variableValue) {
                    $foundNodes = $targetXPath->query('//a:p[a:r/a:t[text()[contains(., "' . $this->templateSymbolStart . $variableKey . $this->templateSymbolEnd . '")]]]');
                    foreach ($foundNodes as $foundNode) {
                        if (isset($options['remove']) && $options['remove']) {
                            // remove paragraph, set by removeVariableText
                            $pChilds = $foundNode->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'p');
                            // if the parent node doesn't include a paragraph node, add it
                            if ($pChilds->length > 1) {
                                $pChilds->item(0)->parentNode->removeChild($foundNode);
                            } else {
                                $emptyP = $foundNode->ownerDocument->createElement('a:p');
                                $foundNode->parentNode->appendChild($emptyP);
                                $foundNode->parentNode->removeChild($foundNode);
                            }
                        } else {
                            // empty paragraph
                            $foundNodesT = $foundNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't');
                            $isFirstNodeT = true;
                            foreach ($foundNodesT as $foundNodeT) {
                                if ($isFirstNodeT) {
                                    // replace only the first occurrence
                                    $foundNodeT->nodeValue = $variableValue;
                                    $isFirstNodeT = false;
                                    continue;
                                } else {
                                    // empty other ocurrences
                                    $foundNodeT->nodeValue = '';
                                }
                            }
                        }
                    }
                }
                $newContent = $targetDOM->saveXML();
            } else {
                // inline replacement
                foreach ($variables as $variableKey => $variableValue) {
                    $variableValue = $this->parseAndCleanTextString($variableValue);
                    $newContent = str_replace($this->templateSymbolStart . $variableKey . $this->templateSymbolEnd, $variableValue, $newContent);
                }
            }

            $this->zipPptx->addContent($targetContent['path'], $newContent);
        }

        PhppptxLogger::logger('Replace variable text.', 'info');
    }

    /**
     * Replaces video placeholders by an external video
     *
     * @access public
     * @param array $variables variable names and video paths
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'image' (array)
     *          'image' image to be used as preview. Set a default one if not set.
     *          'mime' (string) forces a mime (image/jpg, image/jpeg, image/png, image/gif, image/bmp, image/webp)
     *          'usePlaceholderImage' (bool) if true, do not change the placeholder image. Default as false
     *      'mime' (string) forces a mime (video/mp4, video/x-msvideo, video/x-ms-wmv, video/unknown)
     *      'target' (string) slides (default), slideLayouts, slideMasters
     * @throws Exception video doesn't exist
     * @throws Exception video format is not supported
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     */
    public function replaceVariableVideo($variables, $options = array())
    {
        // default values
        if (!isset($options['target'])) {
            $options['target'] = 'slides';
        }

        $targetContents = $this->zipPptx->getContentByType($options['target']);
        if ($options['target'] == 'slides' && isset($options['activeSlide']) && $options['activeSlide']) {
            $targetContents = array($targetContents[$this->activeSlide['position']]);
        }

        foreach ($targetContents as $targetContent) {
            $targetDOM = $this->xmlUtilities->generateDomDocument($targetContent['content']);

            // add the new relationship
            $targetRelsPath = str_replace('ppt/'.$options['target'].'/', 'ppt/'.$options['target'].'/_rels/', $targetContent['path']) . '.rels';
            $targetRelsDOM = $this->zipPptx->getContent($targetRelsPath, 'DOMDocument');

            foreach ($variables as $variableKey => $variableSrc) {
                $newRels = $this->replaceMedia($variableKey, $variableSrc, $targetDOM, 'video', $options);

                if (count($newRels) > 0) {
                    foreach ($newRels as $newRel) {
                        // add the new relationship
                        $this->generateRelationship($targetRelsDOM, $newRel);
                    }
                }
            }

            // refresh contents
            $this->zipPptx->addContent($targetContent['path'], $targetDOM->saveXML());
            if ($targetRelsDOM) {
                $this->zipPptx->addContent($targetRelsPath, $targetRelsDOM->saveXML());
            }

            // free DOMDocument resources
            $targetDOM = null;
            $targetRelsDOM = null;
        }

        PhppptxLogger::logger('Replace variable video.', 'info');
    }

    /**
     * Extract the variables from a string
     *
     * @access private
     * @param string $content
     * @return array $variables
     */
    private function extractVariables($content) {
        $matches = array();
        preg_match_all('/'.preg_quote($this->templateSymbolStart, '/').self::$regExprVariableSymbols.preg_quote($this->templateSymbolEnd, '/').'/msiU', $content, $matches);

        $variables = array();
        foreach ($matches[0] as $variable) {
            $variables[] = str_replace(array($this->templateSymbolStart, $this->templateSymbolEnd), '', $variable);
        }

        return $variables;
    }

    /**
     * Removes placeholder media
     *
     * @access public
     * @param string $variable variable to remove
     * @param DOMDocument $domContent
     * @param DOMXPath $xPathContent
     * @param string $type audio, image, video
     * @param array $options
     *      'activeSlide' (bool) if true, get only the active slide (slides target). Default as false
     *      'target' (string) slides (default), slideLayouts, slideMasters
     */
    private function removeMedia($variable, $domContent, $xPathContent, $type, $options = array())
    {
        $nodesVariable = $xPathContent->query('//p:pic[.//p:cNvPr[@descr="'.$this->templateSymbolStart . $variable . $this->templateSymbolEnd.'" or @title="'.$this->templateSymbolStart . $variable . $this->templateSymbolEnd.'"]]');
        foreach ($nodesVariable as $nodeVariable) {
            $nodesMedia = null;
            switch ($type) {
                case 'audio':
                    $nodesMedia = $nodeVariable->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'audioFile');
                    break;
                case 'image':
                    $nodesMedia = $nodeVariable->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip');
                    break;
                case 'video':
                    $nodesMedia = $nodeVariable->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'videoFile');
                    break;
            }
            // check if it's the correct media type
            if ($nodesMedia && $nodesMedia->length > 0) {
                // remove timing related tags
                $nodesCNvPr = $nodeVariable->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cNvPr');
                if ($nodesCNvPr->length > 0) {
                    $idCNvPr = $nodesCNvPr->item(0)->getAttribute('id');
                    if ($idCNvPr) {
                        // clean related timing tags
                        $this->cleanTimingTags($idCNvPr, $xPathContent, $type);
                    }
                }

                $nodeVariable->parentNode->removeChild($nodeVariable);
            }
        }
    }

    /**
     * Repairs variables
     *
     * @access private
     */
    private function repairVariables()
    {
        // get the placeholder symbols to parse them
        if (function_exists('mb_str_split')) {
            $templateSymbolStartContents = mb_str_split($this->templateSymbolStart);
            $templateSymbolEndContents = mb_str_split($this->templateSymbolEnd);
        } else {
            if (function_exists('mb_strlen')) {
                $length = mb_strlen($this->templateSymbolStart);
                $templateSymbolStartContents = array();
                for ($i = 0; $i < $length; $i++) {
                    $templateSymbolStartContents[] = mb_substr($this->templateSymbolStart, $i, 1);
                }
                $length = mb_strlen($this->templateSymbolEnd);
                $templateSymbolEndContents = array();
                for ($i = 0; $i < $length; $i++) {
                    $templateSymbolEndContents[] = mb_substr($this->templateSymbolEnd, $i, 1);
                }
            } else {
                $templateSymbolStartContents = str_split($this->templateSymbolStart);
                $templateSymbolEndContents = str_split($this->templateSymbolEnd);
            }
        }
        $templateSymbolContents = array_merge($templateSymbolStartContents, $templateSymbolEndContents);
        // get XPath to query paragraphs that contain placeholder symbols
        $templateSymbolXPathQuery = '//a:p[contains(., "'.array_shift($templateSymbolContents).'")';
        foreach ($templateSymbolContents as $templateSymbolContent) {
            $templateSymbolXPathQuery .= ' and contains(., "'.$templateSymbolContent.'")';
        }
        $templateSymbolXPathQuery .= ']';

        $targetTypes = array('notesSlides', 'slides', 'slideLayouts', 'slideMasters');
        foreach ($targetTypes as $targetType) {
            $targetContents = $this->zipPptx->getContentByType($targetType);
            foreach ($targetContents as $targetContent) {
                $targetNewContent = $targetContent['content'];
                $targetDOM = $this->xmlUtilities->generateDomDocument($targetNewContent);
                $targetXPath = new DOMXPath($targetDOM);
                $targetXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

                $foundPlaceholderNodes = $targetXPath->query($templateSymbolXPathQuery);
                foreach ($foundPlaceholderNodes as $node) {
                    $templateSymbolContents = array_merge($templateSymbolStartContents, $templateSymbolEndContents);
                    $templateSymbolContentQuoted = array();
                    foreach ($templateSymbolContents as $templateSymbolContent) {
                        $templateSymbolContentQuoted[] = preg_quote($templateSymbolContent, '/');
                    }
                    $matchesPlaceholders = array();
                    $templateSymbolXPathRegExpr = implode('(.*)', $templateSymbolContentQuoted);
                    preg_match_all('/'.$templateSymbolXPathRegExpr.'/msiU', $node->ownerDocument->saveXML($node), $matchesPlaceholders);

                    foreach ($matchesPlaceholders[0] as $matchPlaceholders) {
                        $targetNewContent = str_replace($matchPlaceholders, strip_tags($matchPlaceholders), $targetNewContent);
                    }
                }
                $this->zipPptx->addContent($targetContent['path'], $targetNewContent);

                // free DOMDocument resources
                $targetDOM = null;
            }
        }

        $this->preprocessed = true;
    }

    /**
     * Replaces placeholder image
     *
     * @access public
     * @param string $variable this variable uniquely identifies the image we want to replace
     * @param string $src path to the substitution image or stream or base64 image
     * @param DOMDocument $domContent
     * @param array $options
     *      'descr' (string) set a descr value
     *      'mime' (string) forces a mime (image/jpg, image/jpeg, image/png, image/gif, image/bmp, image/webp)
     *      'sizeX' (int) EMUs (English Metric Unit)
     *      'sizeY' (int) EMUs (English Metric Unit)
     *      'target' (string) slides
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     * @throws Exception mime option is not set and getimagesizefromstring is not available
     * @return string new rels
     */
    private function replaceImage($variable, $src, $domContent, $options = array())
    {
        // get image information
        $imageInformation = new ImageUtilities();
        $imageContents = $imageInformation->returnImageContents($src, $options);

        $imageContent = $imageContents['content'];
        $extension = $imageContents['extension'];
        $mimeType = $imageContents['mime'];

        // fill this variable only if the placeholder is found
        $relString = '';

        // transitional mode
        $xPathContent = new DOMXPath($domContent);
        $xPathContent->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $xPathContent->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        $domImagesPlaceholder = $xPathContent->query('//p:pic[.//p:cNvPr[@descr="'.$this->templateSymbolStart . $variable . $this->templateSymbolEnd.'" or @title="'.$this->templateSymbolStart . $variable . $this->templateSymbolEnd.'"]]');
        foreach ($domImagesPlaceholder as $domImagePlaceholder) {
            $nodesBlipFill = $domImagePlaceholder->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'blipFill');
            if ($nodesBlipFill->length > 0) {
                // create a new Id
                $idImage = $this->generateUniqueId();
                $ridImage = 'rId' . $idImage;
                // generate the new relationship
                $relString = '<Relationship Id="' . $ridImage . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img' . $idImage . '.' . $extension . '" />';
                // generate content type if it does not exist yet
                $this->generateDefault($extension, 'image/' . $extension);

                $nodesBlipFill->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')->item(0)->setAttribute('r:embed', $ridImage);

                // apply custom sizes
                if (isset($options['sizeX']) || isset($options['sizeY'])) {
                    $nodesExt = $xPathContent->query('.//p:spPr/a:xfrm/a:ext', $domImagePlaceholder);
                    if ($nodesExt->length > 0) {
                        if (isset($options['sizeX'])) {
                            $nodesExt->item(0)->setAttribute('cx', $options['sizeX']);
                        }
                        if (isset($options['sizeY'])) {
                            $nodesExt->item(0)->setAttribute('cy', $options['sizeY']);
                        }
                    }
                }

                // apply custom descr
                if (isset($options['descr'])) {
                    $nodesCNvPr = $domImagePlaceholder->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cNvPr');
                    if ($nodesCNvPr->length > 0) {
                        $nodesCNvPr->item(0)->setAttribute('descr', $options['descr']);
                    }
                }

                // copy the image in the template with the new name
                $this->zipPptx->addContent('ppt/media/img' . $idImage . '.' . $extension, $imageContent);
            }
        }

        return $relString;
    }

    /**
     * Replaces list values in a recursive way
     *
     * @access public
     * @param string $search
     * @param DOMNode $domNode
     * @param array $listValues
     * @param int $level
     * @param array $options
     * @param array $slidesContent
     */
    private function replaceListValues($search, $domNode, $listValues, $level = 0, $options = array(), $slidesContent = array()) {
        foreach ($listValues as $key => $value) {
            if (is_array($value)) {
                $this->replaceListValues($search, $domNode, $value, $level + 1, $options);
            } else {
                $newNode = $domNode->cloneNode(true);
                $textNodes = $newNode->getElementsBytagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't');
                foreach ($textNodes as $text) {
                    if ($value instanceof PptxFragment) {
                        // PptxFragment replacement

                        // remove existing a:r tags
                        $rNodes = $newNode->getElementsBytagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'r');
                        $nodesToBeRemoved = array();
                        foreach ($rNodes as $rNode) {
                            $nodesToBeRemoved[] = $rNode;
                        }
                        foreach ($nodesToBeRemoved as $nodeToBeRemoved) {
                            $nodeToBeRemoved->parentNode->removeChild($nodeToBeRemoved);
                        }

                        // import the new contents
                        $newNodesFragment = $value->inlinePptxXML();

                        foreach ($newNodesFragment as $newNodeFragment) {
                            $newContentFragment = $newNode->ownerDocument->createDocumentFragment();
                            $newContentFragment->appendXML($newNodeFragment->ownerDocument->saveXML($newNodeFragment));
                            $newNode->appendChild($newContentFragment);
                        }

                        // handle external relationships such as hyperlinks
                        $externalRelationships = $value->getExternalRelationships();
                        if (count($externalRelationships) > 0) {
                            $slideRelsPath = str_replace('ppt/'.$options['target'].'/', 'ppt/'.$options['target'].'/_rels/', $slidesContent['path']) . '.rels';
                            $slideRelsDOM = $this->zipPptx->getContent($slideRelsPath, 'DOMDocument');

                            $this->addExternalRelationships($externalRelationships, $slidesContent['path'], $slideRelsDOM);

                            // refresh contents
                            $this->zipPptx->addContent($slideRelsPath, $slideRelsDOM->saveXML());

                            // free DOMDocument resources
                            $slideRelsDOM = null;
                        }
                    } else {
                        // text replacement

                        $sxText = simplexml_import_dom($text);
                        $strNode = (string)$sxText;
                        $strNodeReplaced = str_replace($search, $value, $strNode);
                        $sxText[0] = $strNodeReplaced;
                    }
                }
                if ($level > 0) {
                    $numPprNodes = $newNode->getElementsBytagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'pPr');
                    // set list level only if greater than 0
                    if ($numPprNodes->length > 0) {
                        // a pPr node exists
                        $numPprNodes->item(0)->setAttribute('lvl', $level);
                    } else {
                        $newPPr = '<a:pPr lvl="'.$level.'" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>';
                        $newPPrFragment = $newNode->ownerDocument->createDocumentFragment();
                        $newPPrFragment->appendXML($newPPr);

                        $numRNodes = $newNode->getElementsBytagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'r');
                        if ($numRNodes->length > 0) {
                            $numRNodes->item(0)->parentNode->insertBefore($newPPrFragment, $numRNodes->item(0));
                        }
                    }
                }
                $domNode->parentNode->insertBefore($newNode, $domNode);
            }
        }
    }

    /**
     * Replaces placeholder media
     *
     * @access public
     * @param string $variable variable to replace
     * @param string $src path to the substitution media
     * @param DOMDocument $domContent
     * @param string $type audio, video
     * @param array $options
     *      'image' (array)
     *          'image' image to be used as preview. Set a default one if not set.
     *          'mime' (string) forces a mime (image/jpg, image/jpeg, image/png, image/gif, image/bmp, image/webp)
     *          'usePlaceholderImage' (bool) if true, do not change the placeholder image. Default as false
     *      'mime' (string) forces a mime
     *      'target' (string) slides (default), slideLayouts, slideMasters
     * @throws Exception media doesn't exist
     * @throws Exception media format is not supported
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     * @throws Exception mime option is not set and getimagesizefromstring is not available
     * @return array new rels
     */
    private function replaceMedia($variable, $src, $domContent, $type, $options = array())
    {
        // get media content
        $mediaContents = array();
        if ($type == 'audio') {
            $audioInformation = new AudioUtilities();
            $mediaContents = $audioInformation->returnAudioContents($src, $options);
        } else if ($type == 'video') {
            $videoInformation = new VideoUtilities();
            $mediaContents = $videoInformation->returnVideoContents($src, $options);
        }
        $mediaStream = $mediaContents['content'];

        // fill this variable only if the placeholder is found
        $relString = array();

        // transitional mode
        $xPathContent = new DOMXPath($domContent);
        $xPathContent->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $xPathContent->registerNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $domMediasPlaceholder = $xPathContent->query('//p:pic[.//p:cNvPr[@descr="'.$this->templateSymbolStart . $variable . $this->templateSymbolEnd.'" or @title="'.$this->templateSymbolStart . $variable . $this->templateSymbolEnd.'"]]');
        $domMedias = $domContent->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'cNvPr');
        foreach ($domMediasPlaceholder as $domMediaPlaceholder) {
            $contentType2006 = '';
            $contentType2007 = '';
            if ($type == 'audio') {
                $nodesMedia = $domMediaPlaceholder->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'audioFile');
                $contentType2006 = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio';
                $contentType2007 = 'http://schemas.microsoft.com/office/2007/relationships/media';
            } else if ($type == 'video') {
                $nodesMedia = $domMediaPlaceholder->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'videoFile');
                $contentType2006 = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video';
                $contentType2007 = 'http://schemas.microsoft.com/office/2007/relationships/media';
            }
            if (isset($nodesMedia) && $nodesMedia->length > 0) {
                // handle media replacement. Two values are needed: 2006 and 2007
                // create a new Id
                $idMedia2006 = $this->generateUniqueId();
                $ridMedia2006 = 'rId' . $idMedia2006;
                $idMedia2007 = $this->generateUniqueId();
                $ridMedia2007 = 'rId' . $idMedia2007;
                // path file
                $idMedia = $this->generateUniqueId();

                // generate the new relationships. Both have the same target
                $relString[] = '<Relationship Id="' . $ridMedia2006 . '" Type="'.$contentType2006.'" Target="../media/'.$type.'' . $idMedia . '.' . $mediaContents['extension'] . '" />';
                $relString[] = '<Relationship Id="' . $ridMedia2007 . '" Type="'.$contentType2007.'" Target="../media/'.$type.'' . $idMedia . '.' . $mediaContents['extension'] . '" />';

                // generate content type if it does not exist yet
                $this->generateDefault($mediaContents['extension'], $mediaContents['mime']);

                // update 2006 rId
                $nodesMedia->item(0)->setAttribute('r:link', $ridMedia2006);
                // update 2007 rId
                $nodesExtLst = $nodesMedia->item(0)->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'extLst');
                if ($nodesExtLst->length > 0) {
                    $nodesExt = $nodesExtLst->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'ext');
                    if ($nodesExt->length > 0) {
                        $nodesP14Media = $nodesExt->item(0)->getElementsByTagNameNS('http://schemas.microsoft.com/office/powerpoint/2010/main', 'media');
                        if ($nodesP14Media->length > 0) {
                            $nodesP14Media->item(0)->setAttribute('r:embed', $ridMedia2007);
                        }
                    }
                }

                // copy the media in the template. Only one file copy for both 2006 and 2007
                $this->zipPptx->addContent('ppt/media/' . $type . $idMedia . '.' . $mediaContents['extension'], $mediaStream);

                // handle image replacement
                $replaceImage = true;
                if (isset($options['image']) && isset($options['image']['usePlaceholderImage']) && $options['image']['usePlaceholderImage']) {
                    // use the existing placeholder image. Do not add a new one
                    $replaceImage = false;
                }
                if ($replaceImage) {
                    // do the image replacement
                    $imageContent = '';
                    if (isset($options['image']) && isset($options['image']['image'])) {
                        $imageInformation = new ImageUtilities();
                        $imageContents = $imageInformation->returnImageContents($options['image']['image'], $options);

                        $imageContent = $imageContents['content'];
                        $extension = $imageContents['extension'];
                        $mimeType = $imageContents['mime'];
                    } else {
                        // default image
                        if ($type == 'audio') {
                            $imageContent = base64_decode('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');
                        } else if ($type == 'video') {
                            $imageContent = base64_decode('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');
                        }
                        $extension = 'png';
                        $mimeType = 'image/png';
                    }

                    // create a new Id
                    $idImage = $this->generateUniqueId();
                    $ridImage = 'rId' . $idImage;
                    // generate the new relationship
                    $relString[] = '<Relationship Id="' . $ridImage . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img' . $idImage . '.' . $extension . '" />';
                    // generate content type if it does not exist yet
                    $this->generateDefault($extension, 'image/' . $extension);

                    $nodesBlipFill = $domMediaPlaceholder->getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'blipFill');
                    if ($nodesBlipFill->length > 0) {
                        $nodesBlip = $nodesBlipFill->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip');
                        if ($nodesBlip->length > 0) {
                            $nodesBlip->item(0)->setAttribute('r:embed', $ridImage);

                            // copy the image in the template with the new name
                            $this->zipPptx->addContent('ppt/media/img' . $idImage . '.' . $extension, $imageContent);
                        }
                    }
                }
            }
        }

        return $relString;
    }
}
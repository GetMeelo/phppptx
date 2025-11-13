<?php
/*
 * phppptx configuration plugin test
 */
require_once 'classes/CreatePptx.php';

$output = '';

$break = isset($_SERVER['HTTP_USER_AGENT']) ? '<br />' : PHP_EOL;
$isWeb = isset($_SERVER['HTTP_USER_AGENT']) ? true : false;

if ($isWeb) {
    $output .= '<html>';
    $output .= '<head>';
    $output .= '<title>phppptx conversion plugin checker</title>';
    $output .=
        '<style type="text/css">
        body {
            background: #F3F3F3;
            font-family: Open Sans, Sans-Serif;
            color: #343434;
        }
        #page {
            border: 1px solid #ababab;
            margin: 15px auto 0 auto;
            width: 900px;
            box-shadow: 5px 10px 10px #bdc3c5;
            background: white;
        }
        #info {
            background: -moz-linear-gradient(top, #555555, #343434) !important;
            background: linear-gradient(#555555, #343434);
            padding: 20px 25px 15px 25px;
        }
        #header {
            background: url("https://www.phppptx.com/files/phpdocx/logo.png") no-repeat;
            height: 100px;
            padding: 0 10px 10px 125px;
        }
        #header h1 {
            margin: 0px;
            color: #fff;
        }
        h1 .grey {
            color: #242424;
        }
        .white {
            color: #fff;
        }
        .color {
            color: #C060E0 !important;
        }
        #sidebar {
            float: left;
            width: 90px;
        }
        .divider {
            border-top: 1px solid #ababab;
            padding: 15px 25px 15px 20px;
            background: #fff !important;
            color: #000;
        }
        ul {
            list-style-type: none;
            margin:20px;
            padding:0;
        }
        li {
            padding:10px;
        }
        li div {
            border-top: 1px solid #EEE;
            margin-left: 65px;
            margin-bottom: 10px;
        }
        .testok {
            padding:5px;
            margin: 0 10px 0 0;
            color: #FFFFFF;
            background-color: #008000;
            -webkit-border-radius: 5px;
            -moz-border-radius: 5px;
            border-radius: 5px;
            display: inline-block;
            width:50px;
            font-size: 11px;
            text-align:center;
        }
        .testko {
            padding:5px;
            margin: 0 10px 0 0;
            color: #FFFFFF;
            background-color: #FE2E2E;
            -webkit-border-radius: 5px;
            -moz-border-radius: 5px;
            border-radius: 5px;
            display: inline-block;
            width:50px;
            font-size: 11px;
            text-align:center;
        }
        .testwarn {
            padding:5px;
            margin: 0 10px 0 0;
            color: #FFFFFF;
            background-color: #dd9118;
            -webkit-border-radius: 5px;
            -moz-border-radius: 5px;
            border-radius: 5px;
            display: inline-block;
            width:50px;
            font-size: 11px;
            text-align:center;
        }
        .comment {
            margin: 70px;
            font-size: 11px;
        }
        textarea {
            width: 650px;
            height: 150px;
            margin-top: 10px;
        }
        .info-list {
            margin: 0px 20px 20px 0px;
        }
        .clear {clear: both;}

        .tab {
            overflow: hidden;
            /*border: 1px solid #ccc;
            background-color: #f1f1f1;
            background: #343434;
            background: -webkit-gradient(linear, left top, left bottom, from(#555555), to(#343434));
            background: -moz-linear-gradient(top, #555555, #343434);
            background: -o-linear-gradient(top, #555555, #343434);*/
        }

        .tab button {
          background-color: inherit;
          float: left;
          border: none;
          outline: none;
          cursor: pointer;
          padding: 14px 16px 14px 0px;
          transition: 0.3s;
          text-transform: uppercase;
          font-weight: bold;
          color: #ffffff;
          font-family: Open Sans, Sans-Serif;
          font-size: 14px;
        }

        .tab .tablinks {
            margin-top: 20px;
            margin-left: 50px;
        }

        .tab button:hover {
          /*background-color: #ddd;*/
          color: #C060E0;
        }

        .tab button.active {
          /*background-color: #ccc;*/
        }

        .tabcontent {
          display: none;
          padding: 6px 12px;
          border: 1px solid #ccc;
          border-top: none;
        }
        .secondList {
            margin: 0px;
            list-style-type: circle;
            padding-left: 35px;
        }

        .introList {margin: 0 auto;}

        .space {margin-bottom: 5px;}

        a {color: #C060E0; text-decoration: none;}

        </style>';
    $output .= '</head>';
    $output .= '<body>';
    $output .= '<div id="page">';

    $output .= '<div id="info">';
    $output .= '<div id="header">';
    $output .= '<h1>php<span class="color">pptx</span></h1>';
    $output .= '<span class="white">Welcome to php<span class="color">pptx</span> conversion plugin checker</span>';
    $output .= '</div>

        <div class="tab">
            <span class="white">Choose one of the following methods. It\'s required to reload the page to refresh the set up changes. </span>
            <button class="tablinks" onclick="openMethod(event, \'LibreOffice\')">LibreOffice</button>
            <button class="tablinks" onclick="openMethod(event, \'MS PowerPoint\')">MS PowerPoint</button>
        </div>


    ';
    $output .= '</div>';

    $output .= '<div>';
    $output .= '<div>';
}

// get phppptxconfig data
$phppptxconfig = PhppptxUtilities::parseConfig();

// libreoffice method
if ($isWeb) {
    $output .= '
        <div id="LibreOffice" class="tabcontent">
            <h3>LibreOffice</h3>
    ';
} else {
    $output .= '-----------' . $break . 'LibreOffice' . $break . '-----------' . $break;
}

if ($isWeb) {
    $output .= '<ul class="checks">';
    $output .= '<li>';
}
if (!isset($phppptxconfig['transform']['method'])) {
    if ($isWeb) {
         $output .= '<span class="testko">';
    }
    $output .= 'Error ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'Please set libreoffice as transform method in config/phppptxconfig.ini' . $break;
} else {
    if ($phppptxconfig['transform']['method'] != 'libreoffice') {
        if ($isWeb) {
         $output .= '<span class="testko">';
        }
        $output .= 'Error ';
        if ($isWeb) {
            $output .= '</span>';
        }
        $output .= 'Please set libreoffice as transform method in config/phppptxconfig.ini' . $break;
    } else {
        if ($isWeb) {
         $output .= '<span class="testok">';
        }
        $output .= 'OK ';
        if ($isWeb) {
            $output .= '</span>';
        }
        $output .= 'libreoffice is set as transform method in config/phppptxconfig.ini' . $break;
    }
}
if ($isWeb) {
    $output .= '</li>';
}

// phppptxconfig path option
if ($isWeb) {
    $output .= '<li>';
}
if (!isset($phppptxconfig['transform']['path'])) {
    if ($isWeb) {
         $output .= '<span class="testko">';
    }
    $output .= 'Error ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'Please set the path to libreoffice main executable in config/phppptxconfig.ini' . $break;
} else {
    if ($isWeb) {
     $output .= '<span class="testok">';
    }
    $output .= 'OK ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'libreoffice path is set in config/phppptxconfig.ini' . $break;
}
if ($isWeb) {
    $output .= '</li>';
}

// phppptxconfig path option exists
if ($isWeb) {
    $output .= '<li>';
}
if (!isset($phppptxconfig['transform']['path']) || !is_readable(str_replace('\'', '', $phppptxconfig['transform']['path'])) || !is_file(str_replace('\'', '', $phppptxconfig['transform']['path']))) {
    if ($isWeb) {
         $output .= '<span class="testko">';
    }
    $output .= 'Error ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'Unable to read libreoffice main executable in the path. Please check the correct path has been added in config/phppptxconfig.ini' . $break;
} else {
    if ($isWeb) {
     $output .= '<span class="testok">';
    }
    $output .= 'OK ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'libreoffice main executable has been found' . $break;
}
if ($isWeb) {
    $output .= '</li>';
}

// get HOME env value
if (isset($phppptxconfig['transform']['home_folder'])) {
    $home_folder = $phppptxconfig['transform']['home_folder'];
} else {
    $home_folder = getenv("HOME");
}

// get rwx HOME folder
if ($isWeb) {
    $output .= '<li>';
}
if (empty($home_folder)) {
    if ($isWeb) {
         $output .= '<span class="testwarn">';
    }
    $output .= 'Warning ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'HOME isn\'t set as environment variable for the current user or it\'s empty. You need to set a custom value in the web server configuration or in config/phppptxconfig.ini or using the homeFolder option available in transform: ' . $break . $break . '<textarea onClick="this.select();">$pptx->transform(\'file.pptx\', \'output.pdf\', \'libreoffice\', array(\'homeFolder\' => \'/folder\'));</textarea>' . $break;
    $output .= 'The chosen HOME folder must exist and include a .config folder with rwx access.' . $break;
} else {
    if ($isWeb) {
     $output .= '<span class="testok">';
    }
    $output .= 'OK ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'HOME environment variable is set for the current user and has the following value: ' . $home_folder . $break;
}
if ($isWeb) {
    $output .= '</li>';
}

// test rw for .config folder
if ($isWeb) {
    $output .= '<li>';
}
if (!is_readable($home_folder . '/.config/libreoffice') || !is_writable($home_folder . '/.config/libreoffice')) {
    if ($isWeb) {
         $output .= '<span class="testwarn">';
    }
    $output .= 'Warning ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'The .config/libreoffice folder in the HOME folder must have rw access.' . $break;
} else {
    if ($isWeb) {
     $output .= '<span class="testok">';
    }
    $output .= 'OK ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'The .config/libreoffice folder in the HOME folder has rw access.' . $break;
}
if ($isWeb) {
    $output .= '</li>';
}

// links to the documentation
if ($isWeb) {
    $output .= '<li>';
}
if ($isWeb) {
    $output .= '<p class="introList">';
}
$output .= 'The following links gather all the available information about the conversion plugin, including installation and set up steps, most common issues and debug: ' . $break;
if ($isWeb) {
    $output .= '</p>';
}

if ($isWeb) {
    $output .= '<ul class="secondList"><li><a href="https://www.phppptx.com/documentation/conversion-plugin" target="_blank">Conversion plugin documentation</a></li></ul>';
} else {
    $output .= 'Conversion plugin [https://www.phppptx.com/documentation/conversion-plugin]' . $break;
}

if ($isWeb) {
    $output .= '<ul class="secondList"><li><a href="https://www.phppptx.com/documentation/conversion-plugin/common-problems-and-possible-errors" target="_blank">Common problems and possible errors</a></li></ul>';
} else {
    $output .= 'Common problems and possible errors [https://www.phppptx.com/documentation/conversion-plugin/common-problems-and-possible-errors]' . $break;
}

if ($isWeb) {
    $output .= '<ul class="secondList"><li><a href="https://www.phppptx.com/documentation/conversion-plugin/debugging-libreoffice" target="_blank">Debugging LibreOffice</a></li></ul>';
} else {
    $output .= 'Debugging LibreOffice [https://www.phppptx.com/documentation/conversion-plugin/debugging-libreoffice]' . $break;
}
if ($isWeb) {
    $output .= '</li>';
}

if ($isWeb) {
	$output .= '</ul>';
	$output .= '</div>';
}

// mspowerpoint method
if ($isWeb) {
    $output .= '
        <div id="MS PowerPoint" class="tabcontent">
            <h3>MS PowerPoint</h3>
    ';
} else {
    $output .= '-------' . $break . 'MS PowerPoint' . $break . '-------' . $break;
}

if ($isWeb) {
    $output .= '<ul class="checks">';
    $output .= '<li>';
}
if (!isset($phppptxconfig['transform']['method'])) {
    if ($isWeb) {
         $output .= '<span class="testko">';
    }
    $output .= 'Error ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'Please set mspowerpoint as transform method in config/phppptxconfig.ini' . $break;
} else {
    if ($phppptxconfig['transform']['method'] != 'mspowerpoint') {
        if ($isWeb) {
         $output .= '<span class="testko">';
        }
        $output .= 'Error ';
        if ($isWeb) {
            $output .= '</span>';
        }
        $output .= 'Please set mspowerpoint as transform method in config/phppptxconfig.ini' . $break;
    } else {
        if ($isWeb) {
         $output .= '<span class="testok">';
        }
        $output .= 'OK ';
        if ($isWeb) {
            $output .= '</span>';
        }
        $output .= 'mspowerpoint is set as transform method in config/phppptxconfig.ini' . $break;
    }
}
if ($isWeb) {
    $output .= '</li>';
}

// COM support
if ($isWeb) {
    $output .= '<li>';
}
if (!class_exists('COM')) {
    if ($isWeb) {
        $output .= '<span class="testko">';
    }
    $output .= 'Error ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'You must install COM support for PHP.' . $break;
} else {
    if ($isWeb) {
        $output .= '<span class="testok">';
    }
    $output .= 'OK ';
    if ($isWeb) {
        $output .= '</span>';
    }
    $output .= 'DOM support is enabled.' . $break;
}
if ($isWeb) {
    $output .= '</li>';
}

// links to the documentation
if ($isWeb) {
    $output .= '<li>';
}
if ($isWeb) {
    $output .= '<p class="introList">';
}
$output .= 'The following pages gather all the information about the conversion plugin, including installation and set up steps, most commo issues and debug:' . $break;

if ($isWeb) {
    $output .= '</p>';
}

if ($isWeb) {
    $output .= '<ul class="secondList"><li><a href="https://www.phppptx.com/documentation/conversion-plugin" target="_blank">Conversion plugin documentation</a></li></ul>';
} else {
    $output .= 'Conversion plugin [https://www.phppptx.com/documentation/conversion-plugin]' . $break;
}

if ($isWeb) {
    $output .= '<ul class="secondList"><li><a href="https://www.phppptx.com/documentation/conversion-plugin/common-problems-and-possible-errors" target="_blank">Common problems and possible errors</a></li></ul>';
} else {
    $output .= 'Common problems and possible errors [https://www.phppptx.com/documentation/conversion-plugin/common-problems-and-possible-errors]' . $break;
}

if ($isWeb) {
    $output .= '</ul>';
    $output .= '</div>';
}

if ($isWeb) {
    $output .= '</ul>';
    $output .= '</div>';
}


if ($isWeb) {
	$output .= '<div class="clear" />';
	$output .= '</div>';
	$output .= '</div>';
}

if ($isWeb) {
    $output .= '
    <script>
        function openMethod(evt, methodName) {
            var i, tabcontent, tablinks;

            tabcontent = document.getElementsByClassName("tabcontent");
            for (i = 0; i < tabcontent.length; i++) {
                tabcontent[i].style.display = "none";
            }

            tablinks = document.getElementsByClassName("tablinks");
            for (i = 0; i < tablinks.length; i++) {
                tablinks[i].className = tablinks[i].className.replace(" active", "");
            }

            document.getElementById(methodName).style.display = "block";
            evt.currentTarget.className += " active";
        }
    </script>';
}
if ($isWeb) {
	$output .= '</body>';
	$output .= '</html>';
}

echo $output;
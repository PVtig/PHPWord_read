<?php

require 'vendor/autoload.php';


$source = __DIR__."/docs/mo.docx";

$objReader = \PhpOffice\PhpWord\IOFactory::createReader('Word2007');

$phpWord = $objReader->load($source);

$data = array();
$body = '';
foreach($phpWord->getSections() as $section) {
    $arrays = $section->getElements();

    foreach($arrays as $e) {
        if(get_class($e) === 'PhpOffice\PhpWord\Element\TextRun'){
            foreach($e->getElements() as $text){

                $font = $text->getFontStyle();
				
				$size = $font->getSize()/10;
				$bold = $font->isBold() ? 'font-weight:700;' :'';
				$color = $font->getColor();
				$fontFamily = $font->getName();
				
				$body .= '<span style="font-size:' . $size . 'em;font-family:' . $fontFamily . '; '.$bold.'; color:#'.$color.'">';
				$body .= $text->getText().'<br /></span>';
				array_push($data, $text->getText());
            }
        }else if(get_class($e) === 'PhpOffice\PhpWord\Element\TextBreak') {
			$body .= '<br />';
		}
		else {
			$body .= $e->getText();
		}
    }    
    break;
}

include 'templ.php';
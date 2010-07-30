<?php
include 'excel_xml.php';
$excel = new excel_xml();

/**
* Add style to your worksheet, it's reference will be "header"
* You add it as an array with the key being the modifier and the value parameter.
* You can add:
*    - size in pt
*    - font like "Georgia"
*    - color in hex for font color
*    - bgcolor in hex for background color
*    - bold as boolean (bold => 1)
*    - italic as boolean
*    - strike as boolean
*/

$header_style = array(
    'bold'       => 1,
    'size'       => '12',
    'color'      => '#FFFFFF',
    'bgcolor'    => '#4F81BD'
);

$excel->add_style('header', $header_style);

$excel->debug();
/**
* Add row and attach the style "header" to it
*/
$excel->add_row(array(
    'Username',
    'First name',
    'Last name'
), 'header');

/**
* Add some rows, if you encapsulate the string inside asterisks,
* they will get bold using the predefined style "bold"
* If you append "|x" where x is a number, that cell will be
* merged with the x following cells
*/
$excel->add_row(array(
    'Anorgan|2'
));

$excel->add_row(array(
    '*Marin*',
    'Crnkovic'
));

/**
* You don't like the arrays, or already have
* some form of csv generating script that uses strings?
* No biggie, just delimit the string with ";" or ","
*/
$excel->add_row('Some number:;12');

/**
* Tell the object to create the worksheet.
* The passed string is the name of the worksheet
*/
$excel->create_worksheet('Users');

/**
* If you invoke the generate method, you will get the
* XML returned or...
*/
$xml = $excel->generate();

/**
* ... you can pass the whole thing for download with
* the passed string as the filename
*/
$excel->download('Download.xml');
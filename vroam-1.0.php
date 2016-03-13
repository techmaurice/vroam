<?php
// VROAM 1.0 
// Very Rudimentary Office Analyser Module

// For quick and dirty identification of 
// Microsoft Office Compound and Access files
// by Maurice de Rooij, 2016

// $testfiles sold separately
$testfiles = Array(
	"../../testfiles/2016D00050.doc",
	"../../testfiles/Onderzoeksrapport_",
	"../../testfiles/corpus/valid.xls",
	"../../testfiles/corpus/MonteCarlo.xls",
	"../../testfiles/corpus/MonteCarlo.meuk",
	"../../testfiles/corpus/Reviews.xls",
	"../../testfiles/corpus/ecdl+paris2001.ppt",
	"../../testfiles/pocos_database_preservation_NANETH.ppt",
	"../../testfiles/corpus/lorem-ipsum.mht",
	"../../testfiles/corpus/Reviews.mdb",
	"../../testfiles/corpus/acc95.mdb",
	"../../testfiles/corpus/acc97.mdb",
);

foreach($testfiles as $file) {
	analyse_msoffice($file);
	}

function analyse_msoffice($file) {
	
	echo "{$file} = ";
	
	$data = file_get_contents($file);

	$compound = "/^\\xd0\\xcf\\x11\\xe0/ims";

	$access = "/^\\x00\\x01\\x00\\x00Standard.{0,1}Jet.{0,1}DB/ims";

	// test if compound header
	$test_compound = preg_match($compound, $data, $matches);

	// test if access header
	$test_access = preg_match($access, $data, $matches);

	if($test_compound > 0) {

		$test = preg_match_all("/(Microsoft.{0,1}PowerPoint|PowerPoint|W.{0,1}o.{0,1}r.{0,1}k.{0,1}b.{0,1}o.{0,1}o.{0,1}k|Microsoft.{0,1}Excel|Microsoft.{0,1}JET|MSWordDoc|Microsoft.{0,1}Word)+/ims", $data, $matches);

		if($test > 0) {

			if(in_array("PowerPoint", $matches[0], true) or in_array("Microsoft PowerPoint", $matches[0], true)) {
				echo "Microsoft PowerPoint\n";
				}

			if(in_array("WorkBook", $matches[0], true) or in_array("Microsoft Excel", $matches[0], true) or in_array("Microsoft JET", $matches[0], true)) {
				echo "Microsoft Excel\n";
				}

			if(in_array("MSWordDoc", $matches[0], true) or in_array("Microsoft Word", $matches[0], true)) {
				echo "Microsoft Word\n";
				}

			}

		}

	else if($test_access > 0) {
		echo "Microsoft Access\n";	
		}

	else {
		echo "Probably not a Microsoft Office document\n";
		}

}

?>
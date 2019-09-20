<?php
require_once('vendor/autoload.php');

use \PhpOffice\PhpSpreadsheet\IOFactory;
use \PhpOffice\PhpSpreadsheet\Writer\IWriter;
use \PhpOffice\PhpSpreadsheet\Reader\IReader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

$currentDate = date('d.m.Y');
$treport = getReport($currentDate);

/* Downloading today's report */
if($treport) {
  $spreadsheetOne = $reader->load("reports/".$treport);
  $todaysReport = $spreadsheetOne->getActiveSheet()->toArray(NULL, true, true, true);
} else {die;}


$yesterday = date('d.m.Y', strtotime('-1 day', strtotime($currentDate)));
$yreport = getReport($yesterday);

/* Downloading yesterday's report */
if($yreport) {
  $spreadsheetTwo = $reader->load("reports/".$yreport);
  $yesterdaysReport = $spreadsheetTwo->getActiveSheet()->toArray(NULL, true, true, true);
} else {
  /* If no report yesterday then search earlier dates */
  $earlierday = 0;
  while (!getReport($yesterday) && $earlierday <= 100) {
    $yesterday = date('d.m.Y', strtotime('-1 day', strtotime($yesterday)));
    $earlierday++;
  }
  $earlierday = 0;

  /* Check if previous report is found */
  if($yreport) {
    $spreadsheetTwo = $reader->load("reports/".$yreport);
    $yesterdaysReport = $spreadsheetTwo->getActiveSheet()->toArray(NULL, true, true, true);
  } else {die;}
}

/* Redeclare yesterday date for record's date, which should be the day before */
$yesterday = date('d.m.Y', strtotime('-1 day', strtotime($currentDate)));

/* Reading through yesterday's report */
$length = count($yesterdaysReport);
for($i = 2; $i<=$length; $i++){

  /* Check if there is a difference with today's report */
  $ammount = checkWithTodays($yesterdaysReport[$i]['A'], $yesterdaysReport[$i]['B'], $todaysReport);
  if($ammount != 0){
    /* If there is create a new record */
    addNewRecord($yesterday, $yesterdaysReport[$i]['A'], $ammount);
  }
}

function checkWithTodays($bankAccount, $ammount, $todaysReport) {

  /* Loop through today's report */
  $length = count($todaysReport);
  for ($i=0; $i <= $length; $i++) {

    if($todaysReport[$i]['A'] == $bankAccount) {

      /* Calculate amount credited */
      $difference = $todaysReport[$i]['B'] - $ammount;
      return $difference;

    }
  }
  return 0;
}

function addNewRecord($date, $bankAccount, $ammount) {

  $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
  $spreadsheet = $reader->load("finaleReport/report.xlsx");
  $report = $spreadsheet->getActiveSheet()->toArray(NULL, true, true, true);

  /* Adding new record */
  $newRecord = array(
    "A" => $date,
    "B" => $bankAccount,
    "C" => $ammount
  );
  array_push($report, $newRecord);
  $newRecord = array();

  /* Updating report */
  $spreadsheet->getActiveSheet()->fromArray($report, NULL, 'A1');
  $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
  $writer->save("finaleReport/report.xlsx");
}

function getReport($date) {

  /* Loop through directory and find files with certain creation date ($date) */
  $dir = new DirectoryIterator(dirname(__FILE__) . "/reports");
  foreach ($dir as $fileinfo) {
      if (!$fileinfo->isDot()) {
          $filename = $fileinfo->getFilename();
          $filedate = filectime('reports/'.$filename);
          $filedate = date('d.m.Y', $filedate);
          if($filedate == $date) {
            return $filename;
          }
      }
  }
  return false;
}

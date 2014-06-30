<?php

/**
 * @file
 * The include file for library functions.
 */

function phpword_test() {

  $phpword = phpword();

  $section = $phpword->addSection();

  $bigBold = new \PhpOffice\PhpWord\Style\Font();
  $bigBold->setBold(true);
  $bigBold->setName('Ubuntu');
  $bigBold->setSize(40);

  // After creating a section, you can append elements:
  $section->addText('Hello Word! :)', $bigBold);

  $B = new \PHPOffice\PhpWord\Style\Font();
  $B->setColor('ff89cf');
  $B->setName('comic sans ms');
  $B->setItalic();
  $B->setSize(16);
  $B->setBgColor('000');

  $section->addText(" This should be a delightful colour of pink! ", $B);

  $created_file = phpword_save('hello-world.docx');

  $message = "Successfully created file $created_file->uri";

  if (function_exists('drush_print_r')) {
    drush_print_r($message);
  }
  else {
    drupal_set_message($message);
  }

  return $phpword;
}

function phpword() {

  // Attempt to load statically cached version.
  static $phpword;
  if (isset($phpword) && is_object($phpword)) {
    return $phpword;
  }

  // Load the library.
  $library = libraries_load('phpword');
  if (empty($library['loaded'])) {
    watchdog('phpword', "Couldn't find the PHPWord library.", array(), WATCHDOG_ERROR);
    drupal_set_message("Couldn't find the PHPWord library.");
    return drupal_exit();
  }

  // Register and load the Autoloader/classes.
  $register = \PHPOffice\PhpWord\Autoloader::register();
  $phpword = new \PHPOffice\PHPWord\PhpWord();

  // Defaults.
  // @todo: Make these customisable.
  $phpword->setDefaultFontName('Arial');
  $phpword->setDefaultFontSize(12);

  return $phpword;

}

function phpword_save($filename = NULL) {

  // Intantiate PhpWord, if it has not already been.
  $phpword = phpword();

  // Create a default filename if not defined.
  if (!$filename) {
    $filename = 'document.docx';
  }
  else {
    // Append .docx if it's not been supplied.
    if (substr($filename, -strlen('.docx')) !== '.docx') {
      $filename .= '.docx';
    }
    // Replace spaces with dashes.
    if (stripos($filename, ' ')) {
      $filename = str_replace(' ', '-', $filename);
    }
  }

  $tmpDir     = drupal_realpath(file_directory_temp());
  $privateDir = 'private://phpword';
  $tmp        = "$tmpDir/$filename";

  // Create the writer and output it to the temp directory.
  $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpword, 'Word2007');
  $objWriter->save($tmp);

  // Create our private phpword directory if it doesn't exist.
  if (!file_prepare_directory($privateDir)) {
    drupal_mkdir($privateDir);
  }

  //Define the output file and save it into file_managed.
  $outputFile = (object) array('uri' => $tmp);
  $fileObject = file_move($outputFile, $privateDir);

  // Extend the file object with the pure filename.
  $fileObject->filename = $filename;

  return $fileObject;

}

function phpword_style() {
  //
}

function phpword_download($file) {

  $filename_full = drupal_realpath($file->uri);
  $filename_stub = $file->filename;

  header('Pragma: ');

  header("Cache-Control: private");
  header("Content-Type: application/docx");

  header("Content-Disposition: $attachment; filename=\"$filename_stub\"");

  echo readfile($filename_full);

  flush();
}

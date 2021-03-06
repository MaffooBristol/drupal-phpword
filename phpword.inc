<?php

/**
 * @file
 * The include file for PHPWord library functions.
 */

/**
 * Load, cache and instantiate PhpWord.
 *
 * @return object
 *   An instantiated and statically cached PhpWord object.
 */
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

  // Instantiate the PhpWord class.
  $phpword = new \PhpOffice\PhpWord\PhpWord();

  // Defaults.
  // @todo: Make these customisable.
  $phpword->setDefaultFontName('Arial');
  $phpword->setDefaultFontSize(12);

  return $phpword;

}

/**
 * Take the cached PhpWord object, save it to temp and move to private/phpword.
 *
 * @param string $filename
 *   The filename, with or without the docx extension. If omitted then a default
 *   value will be used.
 * @param string $directory
 *   The directory in which the file will be saved.
 *
 * @return object
 *   The generated Drupal managed file object.
 */
function phpword_save($filename = NULL, $directory = 'private://phpword') {

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

  $tmpDir = drupal_realpath(file_directory_temp());
  $tmp    = "$tmpDir/$filename";

  // Create the writer and output it to the temp directory.
  $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpword, 'Word2007');
  $objWriter->save($tmp);

  // Create our final phpword directory if it doesn't exist.
  if (!file_prepare_directory($directory)) {
    drupal_mkdir($directory);
  }

  // Normalise trailing slashes.
  $directory = $directory . (substr($directory, -1) !== '/' ? '/' : '');

  //Define the output file and save it into file_managed.
  $outputFile = (object) array(
    'uri' => $tmp,
    'uid' => $GLOBALS['user']->uid,
  );
  $fileObject = file_move($outputFile, $directory);

  // Extend the file object with the pure filename.
  $fileObject->filename = $filename;

  return $fileObject;

}

/**
 * Given a file path, make the file automatically download by setting headers.
 *
 * @param object $file
 *   A managed file object, such as that generated by phpword_save().
 */
function phpword_download($file = NULL) {

  if (!$file) {
    watchdog('phpword', 'No file supplied to phpword_download().', array(), WATCHDOG_WARNING);
    return FALSE;
  }

  $filename_full = drupal_realpath($file->uri);
  $filename_stub = $file->filename;

  drupal_add_http_header("Pragma", " ");

  drupal_add_http_header("Cache-Control", "private");
  drupal_add_http_header("Content-Type", "application/docx");

  drupal_add_http_header("Content-Disposition", "attachment; filename=\"$filename_stub\"");

  echo readfile($filename_full);

  flush();
}

/**
 * Gets the full path to the file and redirects the browser to it.
 *
 * @param object $file
 *   A managed file object, such as that generated by phpword_save().
 */
function phpword_download_direct_path($file = NULL) {

  drupal_goto(file_create_url($file->uri));

  drupal_exit();
}

/**
 * Convert a PNG to JPEG, converting transparency to white.
 *
 * @param string $filename
 *   The filename of the image to convert.
 *
 * @return string
 *   The filename of the converted image.
 */
function phpword_convert_png_to_jpg($filename = NULL) {

  if (!isset($filename)) {
    return FALSE;
  }

  $tmpDir  = drupal_realpath(file_directory_temp());
  $tmp     = "$tmpDir/converted.jpg";
  $path    = substr($filename, stripos($filename, 'sites'));
  $path    = drupal_realpath($path);
  $quality = 80;

  $image = imagecreatefrompng($path);
  $bg    = imagecreatetruecolor(imagesx($image), imagesy($image));

  imagefill($bg, 0, 0, imagecolorallocate($bg, 255, 255, 255));
  imagealphablending($bg, TRUE);
  imagecopy($bg, $image, 0, 0, 0, 0, imagesx($image), imagesy($image));
  imagedestroy($image);

  $save = imagejpeg($bg, $tmp, $quality);
  ImageDestroy($bg);

  return $tmp;

}

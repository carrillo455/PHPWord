<?php
  if (!isset($_POST['url'])) {
    echo json_encode(array(
        'data' => null,
        'msg' => 'Acceso no autorizado.'
      )
    );
    exit;
  }

  require_once 'bootstrap.php';

  use PhpOffice\PhpWord\Settings;

  date_default_timezone_set('UTC');
  error_reporting(E_ALL);
  define('CLI', (PHP_SAPI == 'cli') ? true : false);
  define('EOL', CLI ? PHP_EOL : '<br />');

  Settings::loadConfig();

  // Turn output escaping on
  Settings::setOutputEscapingEnabled(true);

  // Return to the caller script when runs by CLI
  if (CLI) {
      return;
  }

  /**
   * Write documents
   *
   * @param \PhpOffice\PhpWord\PhpWord $phpWord
   * @param string $filename
   * @param array $writers
   *
   * @return string
   */
  function write($phpWord, $filename, $targetDir, $writers)
  {
    // $result = '';
    $result = true;

    // Write documents
    foreach ($writers as $format => $extension) {
      // $result .= date('H:i:s') . " Write to {$format} format";
      if (null !== $extension) {
        $targetFile = "{$filename}.{$extension}";
        $target = $targetDir . $targetFile;
        $phpWord->save($target, $format);
      } else {
        // $result .= ' ... NOT DONE!';
        $result = false;
      }
      // $result .= EOL;
    }

    // $result .= getEndingNotes($writers);

    return $result ? $targetFile : $result;
  }

  $backward = '../../';
  $folder = 'files/docs/';
  $basename =  basename($_POST['url']);
  $source = $backward . $folder . $basename;
  $phpWord = \PhpOffice\PhpWord\IOFactory::load($source);
  // echo $phpWord;

  // Save file
  $folder = 'files/html/';
  $filename = uniqid();
  $targetDir = $backward . $folder;
  $writers = array('HTML' => 'html');
  echo json_encode(
    array(
      'data' => write($phpWord, $filename, $targetDir, $writers)
    )
  );
?>
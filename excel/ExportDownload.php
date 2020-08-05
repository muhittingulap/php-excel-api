<?PHP

header('Content-Disposition: attachment;file="' . $_GET["file"] . '";filename="'.$_GET["name"].'"');
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Length: ' . filesize($_GET["file"]));
header('Content-Transfer-Encoding: binary');
header('Cache-Control: must-revalidate');
header('Pragma: public');
readfile($_GET["file"]);
?>
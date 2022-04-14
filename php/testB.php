<?php
$servername = "localhost";
$username = "angelo";
$password = "Angel1996";
$dbname = "test";
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();


$spreadsheet = $reader->load("company.xlsx");

$d=$spreadsheet->getSheet(0)->toArray();

echo count($d);

$sheetData = $spreadsheet->getActiveSheet()->toArray();

// Create connection
$conn = new mysqli($servername, $username, $password, $dbname);

// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
  }

$i=1;
unset($sheetData[0]);

foreach ($sheetData as $t) {
 // process element here;
    $sql = "INSERT INTO company(COMPANY_ID, COMPANY_NAME, COMPANY_CITY )
    VALUES (".$t[0].", '".$t[1]."', '".$t[2]."')";
    
    if ($conn->query($sql) === TRUE) {
    echo "New record created successfully";
    } else {
    echo "Error: " . $sql . "<br>" . $conn->error;
    }
	echo $i."---".$t[0].",".$t[1].",".$t[2]." <br>";
	$i++;
}

$conn->close();
?>
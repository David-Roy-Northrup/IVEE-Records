<?php
require 'vendor/autoload.php';

use Google\Cloud\Firestore\FirestoreClient;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;

// Initialize Firestore
$firestore = new FirestoreClient([
    'projectId' => 'ivee-a3512'
]);

// Get students
$studentsSnap = $firestore->collection('students')->documents();

// Prepare dynamic headers
$attendanceEvents = [];
$collectionPurposes = [];
$documentTypes = [];

// Fetch unique headers
$attendanceDocs = $firestore->collection('attendanceLog')->documents();
foreach ($attendanceDocs as $doc) {
    $event = $doc['eventName'];
    if (!in_array($event, $attendanceEvents)) {
        $attendanceEvents[] = $event;
    }
}

$collectionDocs = $firestore->collection('collectionLog')->documents();
foreach ($collectionDocs as $doc) {
    $purpose = $doc['purpose'];
    if (!in_array($purpose, $collectionPurposes)) {
        $collectionPurposes[] = $purpose;
    }
}

$documentDocs = $firestore->collection('documentLog')->documents();
foreach ($documentDocs as $doc) {
    $type = $doc['documentType'];
    if (!in_array($type, $documentTypes)) {
        $documentTypes[] = $type;
    }
}

// Create Spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Headers
$headers = array_merge(
    ['Name', 'ID'],
    $attendanceEvents,
    $collectionPurposes,
    $documentTypes
);

$col = 1;
foreach ($headers as $header) {
    $sheet->setCellValueByColumnAndRow($col, 1, $header);
    $col++;
}

// Fill student data
$row = 2;
foreach ($studentsSnap as $studentDoc) {
    $studentID = $studentDoc['idNumber'];
    $studentName = $studentDoc['studentName'];

    $sheet->setCellValue("A$row", $studentName);
    $sheet->setCellValue("B$row", $studentID);

    // Attendance
    $col = 3;
    foreach ($attendanceEvents as $event) {
        $attendanceCheck = $firestore->collection('attendanceLog')
            ->where('studentID', '=', $studentID)
            ->where('eventName', '=', $event)
            ->documents();

        $sheet->setCellValueByColumnAndRow($col, $row, $attendanceCheck->isEmpty() ? '' : '✔');
        $col++;
    }

    // Collections
    foreach ($collectionPurposes as $purpose) {
        $collectionCheck = $firestore->collection('collectionLog')
            ->where('studentID', '=', $studentID)
            ->where('purpose', '=', $purpose)
            ->documents();

        if ($collectionCheck->isEmpty()) {
            $sheet->setCellValueByColumnAndRow($col, $row, '');
        } else {
            $sheet->setCellValueByColumnAndRow($col, $row, $collectionCheck->rows()[0]['amount'] ?? '1');
        }
        $col++;
    }

    // Documents
    foreach ($documentTypes as $docType) {
        $docCheck = $firestore->collection('documentLog')
            ->where('studentID', '=', $studentID)
            ->where('documentType', '=', $docType)
            ->documents();

        $sheet->setCellValueByColumnAndRow($col, $row, $docCheck->isEmpty() ? '' : '✔');
        $col++;
    }

    $row++;
}

// Download file
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="student_logs.xls"');
header('Cache-Control: max-age=0');

$writer = new Xls($spreadsheet);
$writer->save('php://output');
exit;

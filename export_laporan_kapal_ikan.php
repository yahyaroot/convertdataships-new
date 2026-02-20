<?php
declare(strict_types=1);

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

require __DIR__ . '/vendor/autoload.php';

/* ================= HEADERS ================= */
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="laporan_kapal_ikan.xlsx"');
header('Cache-Control: max-age=0');

/* ================= LOAD TEMPLATE ================= */
$templatePath = __DIR__ . '/templates/laporan_template3.xlsx';
if (!file_exists($templatePath)) {
    exit('Template Excel tidak ditemukan');
}

$spreadsheet = IOFactory::load($templatePath);
$sheet = $spreadsheet->getActiveSheet();

/* ================= READ JSON ================= */
$data = json_decode(file_get_contents('php://input'), true);
if (!is_array($data)) {
    exit('JSON tidak valid');
}

/* ================= START ROW ================= */
$startRow   = 19; // SESUAI TEMPLATE
$currentRow = $startRow;
$no = 1;

/* ================= STYLE FUNCTION ================= */
$applyStyle = function ($cell, $bgColor) use ($sheet) {
    $sheet->getStyle($cell)->applyFromArray([
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => ['rgb' => $bgColor],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['rgb' => '000000'],
            ],
        ],
    ]);
};

/* ================= WRITE DATA ================= */
foreach ($data as $row) {

    if (empty($row['mmsi']) && empty($row['ship_name'])) {
        continue;
    }

    $sheet->setCellValue("A{$currentRow}", $no++);
    $sheet->setCellValue("B{$currentRow}", $row['time'] ?? '15:00');
    $sheet->setCellValue("C{$currentRow}", $row['ampm'] ?? 'PM');
    $sheet->setCellValue("D{$currentRow}", $row['name'] ?? '');
    $sheet->setCellValue("E{$currentRow}", $row['callsign'] ?? '');
    $sheet->setCellValue("F{$currentRow}", $row['mmsi'] ?? '');

    /* COURSE */
    $sheet->setCellValue("G{$currentRow}", (float)($row['course'] ?? 0));
    $sheet->getStyle("G{$currentRow}")
          ->getNumberFormat()
          ->setFormatCode('000"°"');

    /* SPEED */
    $sheet->setCellValue("H{$currentRow}", (float)($row['speed'] ?? 0));
    $sheet->getStyle("H{$currentRow}")
          ->getNumberFormat()
          ->setFormatCode('0.00" kts"');

    $sheet->setCellValue("I{$currentRow}", $row['imo'] ?? 0);
    $sheet->setCellValue("J{$currentRow}", $row['flag'] ?? 'Indonesia');
    $sheet->setCellValue("K{$currentRow}", $row['type_kapal'] ?? 'Fishing');
    $sheet->setCellValue("L{$currentRow}", $row['destination'] ?? '');

    /* ================= LATITUDE ================= */
    $sheet->setCellValue("M{$currentRow}", (float)($row['latitude1'] ?? 0));
    $sheet->setCellValue("N{$currentRow}", $row['latitude2'] ?? '');
    $sheet->setCellValue("O{$currentRow}", $row['latitude3'] ?? '');
    $sheet->setCellValue("P{$currentRow}", $row['latitude4'] ?? '');

    $sheet->getStyle("M{$currentRow}")
          ->getNumberFormat()
          ->setFormatCode('00"°"');

    /* ================= LONGITUDE ================= */
    $sheet->setCellValue("Q{$currentRow}", (float)($row['longitude1'] ?? 0));
    $sheet->setCellValue("R{$currentRow}", $row['longitude2'] ?? '');
    $sheet->setCellValue("S{$currentRow}", $row['longitude3'] ?? '');

    /* STATUS */
    $sheet->setCellValue("T{$currentRow}", $row['status'] ?? '');

    /* ================= BASE COLORS ================= */
    $applyStyle("A{$currentRow}", 'FFFFFF');
    $applyStyle("B{$currentRow}", 'C0DAF3');
    $applyStyle("C{$currentRow}", 'C0DAF3');
    $applyStyle("D{$currentRow}", 'FFE8D7');
    $applyStyle("E{$currentRow}", 'FFE8D7');
    $applyStyle("F{$currentRow}", 'C0DAF3');
    $applyStyle("G{$currentRow}", 'C0DAF3');
    $applyStyle("H{$currentRow}", 'FFE8D7');
    $applyStyle("I{$currentRow}", 'C0DAF3');
    $applyStyle("J{$currentRow}", 'C0DAF3');
    $applyStyle("K{$currentRow}", 'FFE8D7');
    $applyStyle("L{$currentRow}", 'FFE8D7');

    // Latitude (Purple)
    foreach (['M','N','O','P'] as $c) {
        $applyStyle("{$c}{$currentRow}", 'E5DFED');
    }

    // Longitude (Pink)
    foreach (['Q','R','S'] as $c) {
        $applyStyle("{$c}{$currentRow}", 'F7DBDA');
    }

    // Status
    $applyStyle("T{$currentRow}", 'C0DAF3');

    $currentRow++;
}
// Tambah baris kosong
$currentRow += 2;

// Weather information
// TINGGI GELOMBANG (A-D merged) dan value di E
$sheet->setCellValue("A{$currentRow}", "TINGGI GELOMBANG");
$sheet->mergeCells("A{$currentRow}:D{$currentRow}");
$sheet->setCellValue("E{$currentRow}", ": SEDANG, 1.9 M");
$currentRow++;

// ARAH ANGIN & KECEPATAN ANGIN (A-D merged) dan value di E
$sheet->setCellValue("A{$currentRow}", "ARAH ANGIN & KECEPATAN ANGIN");
$sheet->mergeCells("A{$currentRow}:D{$currentRow}");
$sheet->setCellValue("E{$currentRow}", ": BARAT - TIMUR, 8 - 16 KNOT");
$currentRow++;

// SUHU UDARA (A-D merged) dan value di E
$sheet->setCellValue("A{$currentRow}", "SUHU UDARA");
$sheet->mergeCells("A{$currentRow}:D{$currentRow}");
$sheet->setCellValue("E{$currentRow}", ": 29-31° C");
$currentRow++;

// CUACA (A-D merged) dan value di E
$sheet->setCellValue("A{$currentRow}", "CUACA");
$sheet->mergeCells("A{$currentRow}:D{$currentRow}");
$sheet->setCellValue("E{$currentRow}", ": BERAWAN TEBAL");
$currentRow++;

/* ================= TANDA TANGAN ================= */
// Tanda tangan dimulai dari baris yang sama dengan CUACA (atau baris setelahnya)
// Posisi di kolom M-S (kanan bawah), merged dan rata tengah
$signRow = $currentRow - 3; // Kembali ke baris pertama weather info untuk align dengan CUACA

// Tanggal dan tempat (kolom M-S merged, rata tengah)
$sheet->setCellValue("M{$signRow}", "Karangasem, 31 Desember 2025");
$sheet->mergeCells("M{$signRow}:S{$signRow}");
$sheet->getStyle("M{$signRow}:S{$signRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$signRow += 2;

// Nama jabatan (kolom M-S merged, rata tengah)
$sheet->setCellValue("M{$signRow}", "Piket AIS Stasiun Bakamla Bali");
$sheet->mergeCells("M{$signRow}:S{$signRow}");
$sheet->getStyle("M{$signRow}:S{$signRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$signRow += 2;

// TTD (kolom M-S merged, rata tengah)
$sheet->setCellValue("M{$signRow}", "TTD");
$sheet->mergeCells("M{$signRow}:S{$signRow}");
$sheet->getStyle("M{$signRow}:S{$signRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$signRow += 2;

// Nama (kolom M-S merged, rata tengah)
$sheet->setCellValue("M{$signRow}", "I NYOMAN GELGEL");
$sheet->mergeCells("M{$signRow}:S{$signRow}");
$sheet->getStyle("M{$signRow}:S{$signRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);


/* ================= OUTPUT ================= */
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');
exit;

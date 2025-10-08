<?php if (!defined('BASEPATH')) exit('No direct script access allowed');

/**
 * Simple Excel export helper (BIFF-style)
 * Compatible with classic XLS openable by Excel.
 *
 * Note: This is a minimal binary writer (old-style XLS).
 * It writes basic labels and numeric cells. It does not
 * implement the full BIFF8 format (styles, column widths, multiple sheets).
 *
 * Usage:
 *   // in controller:
 *   $this->load->helper('exportexcel');
 *   $filename = autoExcelFilename('my-export');
 *   export_xls_headers($filename);
 *   xlsBOF();
 *   xlsWriteLabel(0, 0, "Header1");
 *   xlsWriteNumber(1, 0, 123);
 *   xlsEOF();
 *   exit;
 */

/* --- Low-level BIFF records --- */

function xlsBOF()
{
    // BOF: beginning of file (BIFF version 2/3/4 common header)
    echo pack("ssssss", 0x809, 0x8, 0x0, 0x10, 0x0, 0x0);
    return;
}

function xlsEOF()
{
    echo pack("ss", 0x0A, 0x00);
    return;
}

/**
 * Write a numeric cell
 * Row, Col: zero-based
 */
function xlsWriteNumber($Row, $Col, $Value)
{
    // record 0x0203 (NUMBER)
    echo pack("sssss", 0x203, 14, $Row, $Col, 0x0);
    echo pack("d", $Value);
    return;
}

/**
 * Write a text label cell.
 * This version expects $Value in ISO-8859-1/Windows-1252 encoding for broad Excel compatibility.
 * For UTF-8 input, use xlsWriteLabelUtf8.
 */
function xlsWriteLabel($Row, $Col, $Value)
{
    $L = strlen($Value);
    echo pack("ssssss", 0x204, 8 + $L, $Row, $Col, 0x0, $L);
    echo $Value;
    return;
}

/* --- Higher-level helpers --- */

/**
 * Send header for XLS file download
 * Usage: export_xls_headers('myfile.xls');
 */
function export_xls_headers($filename = 'export.xls')
{
    if (!headers_sent()) {
        // Force download dialog
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
        header("Content-Type: application/vnd.ms-excel; charset=UTF-8");
        header("Content-Disposition: attachment; filename=\"$filename\"");
        header("Content-Transfer-Encoding: binary");
    }
}

/**
 * Convert UTF-8 string to a legacy single-byte encoding suitable for basic XLS export.
 * We try a couple of conversions and fallback to removing invalid bytes.
 */
function xls_convert_utf8_to_singlebyte($string)
{
    if ($string === null) return '';

    // If iconv is available, try to convert to Windows-1252 (SQL/Excel-friendly)
    if (function_exists('iconv')) {
        // TRANSLIT attempts substitution for unknown chars; ignore if fails
        $converted = @iconv('UTF-8', 'Windows-1252//TRANSLIT', $string);
        if ($converted !== false) return $converted;
        $converted = @iconv('UTF-8', 'ISO-8859-1//TRANSLIT', $string);
        if ($converted !== false) return $converted;
    }

    // Fallback: strip non-ascii
    return preg_replace('/[^\x20-\x7E]/', '?', $string);
}

/**
 * Write label with UTF-8 input safely converted to single-byte encoding for Excel compatibility.
 */
function xlsWriteLabelUtf8($Row, $Col, $Value)
{
    $val = xls_convert_utf8_to_singlebyte($Value);
    return xlsWriteLabel($Row, $Col, $val);
}

/**
 * Write a date/timestamp into a cell.
 * $value can be:
 *   - integer timestamp (seconds since epoch)
 *   - string parseable via strtotime
 *   - PHP DateTime object
 *
 * We write as an Excel serial number (days since 1899-12-30).
 * Excel default formatting may show as a number; user can apply date format when opening.
 */
function xlsWriteDate($Row, $Col, $Value)
{
    if ($Value instanceof DateTime) {
        $ts = $Value->getTimestamp();
    } elseif (is_numeric($Value)) {
        // assume unix timestamp in seconds
        $ts = (int)$Value;
    } else {
        $ts = strtotime($Value);
    }

    if ($ts === false || $ts <= 0) {
        // invalid date — write as text
        return xlsWriteLabelUtf8($Row, $Col, (string)$Value);
    }

    // Excel's serial date: days since 1899-12-30 including fractional day from time
    // Unix epoch to Excel: (Unix timestamp / 86400) + 25569
    $excelSerial = ($ts / 86400.0) + 25569;

    // write as number (Excel will interpret it as date if user chooses formatting)
    return xlsWriteNumber($Row, $Col, $excelSerial);
}

/**
 * Write an Excel formula into a cell. The string formula should not begin with '=' (we add it).
 * Note: This writes a text record with '=' prefix which Excel will evaluate on load.
 */
function xlsWriteFormula($Row, $Col, $Formula)
{
    // Minimal approach: write formula as label starting with '='.
    // Note: Proper BIFF FORMULA record is more complex; Excel will accept label starting with '=' and evaluate.
    if (substr($Formula, 0, 1) !== '=') {
        $Formula = '=' . $Formula;
    }
    return xlsWriteLabelUtf8($Row, $Col, $Formula);
}

/**
 * Convenience: write a one-dimensional array as a row starting at $startCol
 * $rowIndex: zero-based Excel row
 * $startCol: zero-based Excel column
 */
function xlsWriteRowFromArray($RowIndex, $startCol, array $row)
{
    $c = $startCol;
    foreach ($row as $cell) {
        if (is_numeric($cell) && !is_string($cell)) {
            xlsWriteNumber($RowIndex, $c, $cell);
        } elseif ($cell instanceof DateTime || strtotime($cell) !== false && !is_numeric($cell)) {
            // If it's a parsable date string, attempt date write
            // We check numeric separately to not treat numbers as dates
            xlsWriteDate($RowIndex, $c, $cell);
        } else {
            xlsWriteLabelUtf8($RowIndex, $c, (string)$cell);
        }
        $c++;
    }
}

/**
 * Write a 2D array of data (rows of values). If $headers is true and $data has string keys,
 * the first row will be the header row built from array keys.
 *
 * $data: array of associative arrays or indexed arrays.
 * $startRow, $startCol: zero-based positions
 */
function xlsWriteArray(array $data, $startRow = 0, $startCol = 0, $headers = true)
{
    $r = $startRow;
    if (empty($data)) return;

    // If associative arrays and headers requested, build header row
    if ($headers) {
        // If the first element is associative array, use its keys as headers
        $first = reset($data);
        if (is_array($first) && array_values($first) !== $first) {
            // associative
            $headerRow = array_keys($first);
            xlsWriteRowFromArray($r, $startCol, $headerRow);
            $r++;
            // write data in dict key order
            foreach ($data as $row) {
                $rowOut = [];
                foreach ($headerRow as $k) {
                    $rowOut[] = isset($row[$k]) ? $row[$k] : '';
                }
                xlsWriteRowFromArray($r, $startCol, $rowOut);
                $r++;
            }
            return;
        }
    }

    // Otherwise treat each element as a numeric-indexed row
    foreach ($data as $row) {
        if (is_array($row)) {
            xlsWriteRowFromArray($r, $startCol, $row);
        } else {
            // scalar row
            xlsWriteRowFromArray($r, $startCol, [$row]);
        }
        $r++;
    }
}

/* --- Utility helpers --- */

/**
 * Generate a filename with .xls extension and optional timestamp
 */
function autoExcelFilename($baseName = 'export', $withTimestamp = true)
{
    $ts = $withTimestamp ? date('Ymd_His') : '';
    $name = trim($baseName) === '' ? 'export' : preg_replace('/[^A-Za-z0-9_\-]/', '_', $baseName);
    return $name . ($ts ? "_{$ts}" : "") . ".xls";
}

/* End of file exportexcel_helper.php */
/* Location: ./application/helpers/exportexcel_helper.php */
/* Please DO NOT modify this information : */
/* Enhanced on " . date('Y-m-d H:i:s') . " */
/*public function download_users_xls()
{
    $this->load->helper('exportexcel');

    $data = [
        ['Name' => 'Alice', 'Email' => 'alice@example.com', 'Joined' => '2020-01-15'],
        ['Name' => 'Bob',   'Email' => 'bob@example.com',   'Joined' => '2021-06-02'],
    ];

    $filename = autoExcelFilename('users');
    export_xls_headers($filename);

    xlsBOF();

    // write array with header row (associative arrays)
    xlsWriteArray($data, 0, 0, true);

    xlsEOF();
    exit;
}
*/

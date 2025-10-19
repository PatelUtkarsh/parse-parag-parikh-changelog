<?php

namespace Spock\PhpParseMutualFund;

class ImprovedFilterRowColumns implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter
{
    private array $targetColumns = [];
    private bool $columnsDetected = false;

    public function readCell(string $columnAddress, int $row, string $worksheetName = ''): bool
    {
        // Read first 5000 rows to ensure we capture all data including GRAND TOTAL
        if ($row >= 1 && $row <= 5000) {
            // If columns haven't been detected yet, read all columns in first few rows
            if (!$this->columnsDetected && $row <= 5) {
                return true;
            }

            // Once we know which columns contain stock names and percentages, filter accordingly
            if ($this->columnsDetected && in_array($columnAddress, $this->targetColumns)) {
                return true;
            }

            // For dynamic detection, we'll read columns B through L in first few rows
            if (
                !$this->columnsDetected &&
                in_array($columnAddress, ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'])
            ) {
                return true;
            }
        }

        return false;
    }

    public function setTargetColumns(array $columns): void
    {
        $this->targetColumns = $columns;
        $this->columnsDetected = true;
    }
}

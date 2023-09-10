<?php
namespace Spock\PhpParseMutualFund;

class MyReadFilter implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter {
	public function readCell( $columnAddress, $row, $worksheetName = '' ) {
		//  Read rows 1 to 7 and columns A to E only
		if ( $row >= 1 && $row <= 150 ) {
			if ( in_array( $columnAddress, [ 'B', 'G' ] ) ) {
				return true;
			}
		}

		return false;
	}
}

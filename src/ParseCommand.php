<?php

namespace Spock\PhpParseMutualFund;

use GuzzleHttp\Exception\GuzzleException;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Helper\Table;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;


class ParseCommand extends Command {

	private OutputInterface $output;

	const SHEET_MAP = [
		'flexi'  => 'PPFCF',
		'liquid' => 'PPFLP',
		'hybrid' => 'PPCHF',
		'tax'    => [ 'PPTSF', 'PPETSF' ],
	];

	private const TARGET_SECTION_HEADERS = [
		'Equity & Equity related' => 'equity_and_equity_related',
		'Arbitrage' => 'arbitrage',
		'(b) Reits' => 'b_reits',
		'Equity & Equity related Foreign Investments' => 'equity_foreign_investments',
		'Certificate of Deposit' => 'certificate_of_deposit',
		'Commercial Paper' => 'commercial_paper',
		'Treasury Bill' => 'treasury_bill',
		'Mutual Fund Units' => 'mutual_fund_units',
		'Reverse Repo / TREPS' => 'reverse_repo_treps',
	];

	private const SECTION_TERMINATOR = "GRAND TOTAL";

	private const NON_DATA_IDENTIFIERS = [
		'Total',
		'Sub Total',
		'Last 3 year',
		'Last 1 year',
		'Last 5 year',
		'Since Inception',
		'Clearing Corporation of India Ltd',
		'TOTAL',
		'Equity & Equity related',
		'Listed / awaiting listing',
		'awaiting listing',
		'(MD ',
		'Tbill',
		'Days',
		'#',
		'(a)',
		'(b)',
		'(c)',
		'(d)',
		'(e)',
		'(a) Listed',
		'(b) Listed',
		'Listed',
		'Unlisted',
		'Foreign'
	];

	public function configure(): void {
		$this->setName( 'parse' )
		     ->setDescription( 'Find diff of mutual fund stock holdings.' )
		     ->setHelp( 'This command help you generate diff of share changes per month for a PPFAS mutual fund' )
		     ->addArgument( 'month-diff-x', InputArgument::OPTIONAL, 'For comparing old sheet as X, As if Y - X.', 2 )
		     ->addOption( 'fund-name', 'f', InputOption::VALUE_OPTIONAL, 'Pick from ' . join( ',', array_keys( self::SHEET_MAP ) ), 'tax' )
		     ->addOption( 'month-diff-y', 'y', InputOption::VALUE_OPTIONAL, 'For comparing new sheet as Y, As if Y - X.', 1 );
	}

	public function execute( InputInterface $input, OutputInterface $output ): int {
		$option       = $input->getOption( 'fund-name' );
		$month_old    = intval( $input->getArgument( 'month-diff-x' ) );
		$month_new    = intval( $input->getOption( 'month-diff-y' ) );
		$this->output = $output;

		$old_path = $this->download_handler( $month_old );
		$new_path = $this->download_handler( $month_new );

		if ( ! $old_path || ! $new_path ) {
			$output->writeln( '<error>Unable to download one or both files.</error>' );
			return Command::FAILURE;
		}

		$output->writeln( '<info>Processing Excel files...</info>' );
		$old_sheet = $this->get_excel_column_map( $old_path, $option );
		$new_sheet = $this->get_excel_column_map( $new_path, $option );

		if (empty($old_sheet) || empty($new_sheet)) {
			$output->writeln( '<error>Unable to extract data from Excel files.</error>' );
			return Command::FAILURE;
		}

		// Note: The execute method needs to be refactored to handle the new sectioned data structure
		// For now, we'll flatten the data to maintain compatibility
		$old_flat = $this->flattenSectionData($old_sheet);
		$new_flat = $this->flattenSectionData($new_sheet);

		// Compare key and value = find diff of % and list.
		$column_diff = [];
		$all_companies = array_unique(array_merge(array_keys($old_flat), array_keys($new_flat)));

		foreach ( $all_companies as $name ) {
			$old_percentage = $old_flat[$name] ?? 0;
			$new_percentage = $new_flat[$name] ?? 0;
			$diff = $new_percentage - $old_percentage;

			// Only include if there's a significant change (more than 0.001%)
			if (abs($diff) > 0.00001) {
				$column_diff[$name] = [
					'diff' => $diff,
					'old' => $old_percentage,
					'new' => $new_percentage
				];
			}
		}

		// Sort by absolute difference (biggest changes first)
		uasort($column_diff, function($a, $b) {
			return abs($b['diff']) <=> abs($a['diff']);
		});

		if (empty($column_diff)) {
			$output->writeln( '<comment>No significant changes found between the two periods.</comment>' );
			return Command::SUCCESS;
		}

		$table = new Table( $output );
		$table->setHeaders( [ 'Company Name', 'Change (%)', 'Previous (%)', 'Current (%)' ] );

		foreach ( $column_diff as $name => $data ) {
			$changeColor = $data['diff'] > 0 ? 'green' : 'red';
			$changeSymbol = $data['diff'] > 0 ? '+' : '';

			$table->addRow( [
				$name,
				"<fg={$changeColor}>{$changeSymbol}" . round( $data['diff'] * 100, 4 ) . '%',
				round( $data['old'] * 100, 4 ) . '%',
				round( $data['new'] * 100, 4 ) . '%'
			] );
		}
		$table->render();

		// Also display section-wise breakdown
		$this->displaySectionWiseComparison($old_sheet, $new_sheet, $output);

		return Command::SUCCESS;
	}

	/**
	 * Temporary helper to flatten sectioned data for backward compatibility
	 */
	private function flattenSectionData(array $sectionData): array {
		$flattened = [];
		foreach ($sectionData as $section => $items) {
			foreach ($items as $item) {
				$flattened[$item['name']] = $item['percent'];
			}
		}
		return $flattened;
	}

	public function download_handler( int $month ): string|false {
		$url = $this->generate_url( $month, 'xlsx' );
		$this->output->isVerbose() && $this->output->writeln( 'Downloading file: ' . $url );
		$new_path = $this->download_spreadsheet_in_temp_dir( $url );
		if ( ! $new_path ) {
			$this->output->isVerbose() && $this->output->writeln( 'Unable to download xlsx file, retrying with xls extension: ' . $url );
			// retry download with xls extension.
			$new_url = $this->generate_url( $month );
			return $this->download_spreadsheet_in_temp_dir( $new_url );
		}

		return $new_path;
	}

	public function get_excel_column_map( string $filename, string $sheetCode, int $try = 0 ): array {
		try {
			// Extract file name extension.
			$inputFileType = \PhpOffice\PhpSpreadsheet\IOFactory::identify( $filename );

			// First pass: detect column structure
			$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader( $inputFileType );
			$reader->setReadDataOnly( true );
			$sheetName = self::SHEET_MAP[ $sheetCode ];
			if ( is_array( $sheetName ) ) {
				$sheetName = $sheetName[$try];
			}
			$reader->setLoadSheetsOnly( [ $sheetName ] );

			// Read without filter first to detect structure
			$spreadsheet = $reader->load( $filename );
			$worksheet = $spreadsheet->getActiveSheet();

			// Detect columns by looking at headers and data patterns
			$nameColumn = null;
			$percentColumn = null;

			// Check first 20 rows to find column headers (increased from 10)
			for ($row = 1; $row <= 20; $row++) {
				$rowData = [];
				for ($col = 'B'; $col <= 'L'; $col++) {
					$cell = $worksheet->getCell($col . $row);
					$value = $cell->getValue();
					$rowData[$col] = $value;
				}

				// Look for headers that indicate stock names and percentages
				foreach ($rowData as $col => $value) {
					if (is_string($value)) {
						$lowerValue = strtolower($value);
						if (strpos($lowerValue, 'company') !== false ||
							strpos($lowerValue, 'name') !== false ||
							strpos($lowerValue, 'security') !== false ||
							strpos($lowerValue, 'instrument') !== false) {
							$nameColumn = $col;
						}
						if (strpos($lowerValue, '%') !== false ||
							strpos($lowerValue, 'percent') !== false ||
							strpos($lowerValue, 'weight') !== false ||
							strpos($lowerValue, 'assets') !== false) {
							$percentColumn = $col;
						}
					}
				}

				// If we haven't found headers, look for data patterns
				if (!$nameColumn || !$percentColumn) {
					foreach ($rowData as $col => $value) {
						if (is_string($value) && strlen(trim($value)) > 3 && !$nameColumn) {
							// Check if this looks like a company name
							if (preg_match('/^[A-Za-z][A-Za-z0-9\s&\.\-\(\),]+$/', trim($value))) {
								$nameColumn = $col;
							}
						}
						if (is_numeric($value) && $value > 0 && $value <= 1 && !$percentColumn) {
							$percentColumn = $col;
						}
					}
				}

				if ($nameColumn && $percentColumn) {
					break;
				}
			}

			// Fallback to default columns if detection failed
			if (!$nameColumn) $nameColumn = 'B';
			if (!$percentColumn) $percentColumn = 'G';

			$this->output->isVerbose() && $this->output->writeln("Detected columns: Name={$nameColumn}, Percent={$percentColumn}");

			// Second pass: read only the detected columns with improved filter
			$filterSubset = new ImprovedFilterRowColumns();
			$filterSubset->setTargetColumns([$nameColumn, $percentColumn]);

			$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader( $inputFileType );
			$reader->setReadDataOnly( true );
			$reader->setLoadSheetsOnly( [ $sheetName ] );
			$reader->setReadFilter( $filterSubset );
			$spreadsheet = $reader->load( $filename );
			$worksheet = $spreadsheet->getActiveSheet();

			// Initialize sectioned data processing
			$sections_data = [];
			$current_section_normalized_key = null;
			$current_section_items = [];

			// Iterate through rows for section-based processing
			foreach ( $worksheet->getRowIterator() as $row_object ) {
				$row_number = $row_object->getRowIndex();

				// Get cell value from Column 'B' (assuming section headers are always there)
				$column_b_value = trim((string)$worksheet->getCell('B' . $row_number)->getValue());

				// Termination Check: If we hit GRAND TOTAL, finish processing
				if (strcasecmp($column_b_value, self::SECTION_TERMINATOR) === 0) {
					if ($current_section_normalized_key !== null && !empty($current_section_items)) {
						$sections_data[$current_section_normalized_key] = $current_section_items;
					}
					$this->output->isVerbose() && $this->output->writeln("Found GRAND TOTAL at row {$row_number}, stopping data extraction");
					break;
				}

				// Section Header Check
				$section_found = false;
				foreach (self::TARGET_SECTION_HEADERS as $display_header => $normalized_key) {
					if (strcasecmp($column_b_value, $display_header) === 0) {
						// Save previous section if it exists
						if ($current_section_normalized_key !== null && !empty($current_section_items)) {
							$sections_data[$current_section_normalized_key] = $current_section_items;
						}

						$current_section_normalized_key = $normalized_key;
						$current_section_items = [];
						$section_found = true;
						$this->output->isVerbose() && $this->output->writeln("Found section header: {$display_header} -> {$normalized_key}");
						break;
					}
				}

				if ($section_found) {
					continue; // This row is a header, not data
				}

				// Data Row Processing (if we're inside a section)
				if ($current_section_normalized_key !== null) {
					// Get potential name and percent values
					$name_value = trim((string)$worksheet->getCell($nameColumn . $row_number)->getValue());
					$percent_value_raw = $worksheet->getCell($percentColumn . $row_number)->getValue();

					// Validate name_value as a data item
					if (empty($name_value) || strlen($name_value) <= 2) {
						continue; // Skip empty or too short names
					}

					// Check against NON_DATA_IDENTIFIERS
					$is_non_data = false;
					foreach (self::NON_DATA_IDENTIFIERS as $identifier) {
						if (stripos($name_value, $identifier) !== false) {
							$is_non_data = true;
							break;
						}
					}

					if ($is_non_data) {
						$this->output->isVerbose() && $this->output->writeln("Skipping non-data row: {$name_value}");
						continue;
					}

					// Validate and Convert percent_value_raw
					$percent_numeric = null;
					$percent_str_cleaned = str_replace('%', '', (string)$percent_value_raw);

					if (is_numeric($percent_str_cleaned)) {
						$value = (float)$percent_str_cleaned;
						if (strpos((string)$percent_value_raw, '%') !== false) {
							// If original value had '%', it's like "8.11%", so convert to 0.0811
							$percent_numeric = $value / 100.0;
						} else {
							// Assumed to be already in decimal form (e.g., 0.0811)
							// Or handle cases where it might be a whole number
							if ($value > 1) {
								// Likely a percentage that needs conversion
								$percent_numeric = $value / 100.0;
							} else {
								$percent_numeric = $value;
							}
						}
					}

					if ($percent_numeric === null || $percent_numeric <= 0) {
						continue; // Skip invalid or zero percentages
					}

					// Add valid data item to current section
					$current_section_items[] = ['name' => $name_value, 'percent' => $percent_numeric];
					$this->output->isVerbose() && $this->output->writeln("Added to {$current_section_normalized_key}: {$name_value} ({$percent_numeric})");
				}
			}

			// After Loop Cleanup: handle case where GRAND TOTAL was missing
			if ($current_section_normalized_key !== null && !empty($current_section_items)) {
				$sections_data[$current_section_normalized_key] = $current_section_items;
			}

			$this->output->isVerbose() && $this->output->writeln("Extracted " . count($sections_data) . " sections");

			// Retry Logic (adapted for sectioned data)
			if (empty($sections_data) && $try < 1 && is_array(self::SHEET_MAP[$sheetCode])) {
				$this->output->isVerbose() && $this->output->writeln('Unable to find any sectioned data in sheet: ' . $sheetName . ' retrying with next sheet name.');
				return $this->get_excel_column_map($filename, $sheetCode, $try + 1);
			}

			return $sections_data;

		} catch (\Exception $e) {
			$this->output->writeln('<error>Error processing Excel file: ' . $e->getMessage() . '</error>');
			return [];
		}
	}

	public function download_spreadsheet_in_temp_dir( string $url ): string|false {
		$userAgents = [
			'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
			'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
		];

		$temp_dir  = sys_get_temp_dir();
		$filename  = basename( $url );
		$temp_file = $temp_dir . DIRECTORY_SEPARATOR . strtok( $filename, '?' );
		if ( file_exists( $temp_file ) ) {
			$this->output->isVerbose() && $this->output->writeln( 'File already exists, using cache: ' . $temp_file );
			return $temp_file;
		}

		$fp = fopen( $temp_file, 'w' );
		if (!$fp) {
			$this->output->writeln('<error>Unable to create temporary file: ' . $temp_file . '</error>');
			return false;
		}

		$client = new \GuzzleHttp\Client([
			'headers' => [
				'User-Agent' => $userAgents[ array_rand( $userAgents ) ],
			],
			'timeout' => 30,
			'connect_timeout' => 10
		]);

		try {
			$response = $client->get( $url, [ 'sink' => $fp ] );
			$http_code = $response->getStatusCode();
		} catch ( GuzzleException $e ) {
			$this->output->writeln('<error>HTTP request failed: ' . $e->getMessage() . '</error>');
			fclose( $fp );
			unlink( $temp_file );
			return false;
		}

		fclose( $fp );

		if ( $http_code !== 200 ) {
			$this->output->writeln('<error>HTTP error code: ' . $http_code . '</error>');
			unlink( $temp_file );
			return false;
		}

		$this->output->isVerbose() && $this->output->writeln( 'Downloaded file: ' . $temp_file );
		return $temp_file;
	}

	public function get_last_date_of_datetime( \DateTime $current_date ): \DateTime {
		return new \DateTime( date( "Y-m-t", $current_date->format( 'U' ) ) );
	}

	public function generate_url( int $month_diff, string $extension = 'xls' ): string {
		// Last month date object.
		$current_date = new \DateTime();
		if ( $month_diff !== 0 ) {
			$current_date->modify( "-{$month_diff} month" );
		}
		$last_day_of_month = $this->get_last_date_of_datetime( $current_date );
		$year              = $last_day_of_month->format( 'Y' );
		$month             = $last_day_of_month->format( 'F' );
		$day               = $last_day_of_month->format( 'd' );

		return sprintf( 'https://amc.ppfas.com/downloads/portfolio-disclosure/%1$s/PPFAS_Monthly_Portfolio_Report_%2$s_%3$s_%1$s.%4$s', $year, $month, $day, $extension );
	}

	/**
	 * Display section-wise comparison between old and new data
	 */
	private function displaySectionWiseComparison(array $oldSections, array $newSections, OutputInterface $output): void {
		$output->writeln("\n<info>Section-wise Breakdown:</info>");

		// Get all sections from both datasets
		$allSections = array_unique(array_merge(array_keys($oldSections), array_keys($newSections)));
		sort($allSections);

		foreach ($allSections as $sectionKey) {
			$oldItems = $oldSections[$sectionKey] ?? [];
			$newItems = $newSections[$sectionKey] ?? [];

			// Create lookup arrays for easier comparison
			$oldLookup = [];
			foreach ($oldItems as $item) {
				$oldLookup[$item['name']] = $item['percent'];
			}

			$newLookup = [];
			foreach ($newItems as $item) {
				$newLookup[$item['name']] = $item['percent'];
			}

			// Get all companies in this section
			$allCompanies = array_unique(array_merge(array_keys($oldLookup), array_keys($newLookup)));

			// Calculate differences for this section
			$sectionDiffs = [];
			foreach ($allCompanies as $company) {
				$oldPercent = $oldLookup[$company] ?? 0;
				$newPercent = $newLookup[$company] ?? 0;
				$diff = $newPercent - $oldPercent;

				if (abs($diff) > 0.00001) { // Only include significant changes
					$sectionDiffs[$company] = [
						'diff' => $diff,
						'old' => $oldPercent,
						'new' => $newPercent
					];
				}
			}

			// Skip section if no changes
			if (empty($sectionDiffs)) {
				continue;
			}

			// Sort by absolute difference
			uasort($sectionDiffs, function($a, $b) {
				return abs($b['diff']) <=> abs($a['diff']);
			});

			// Display section header with readable name
			$sectionDisplayName = $this->getSectionDisplayName($sectionKey);

			// Create table for this section with section name in the header
			$sectionTable = new Table($output);
			$sectionTable->setHeaders([$sectionDisplayName, 'Change (%)', 'Previous (%)', 'Current (%)']);

			foreach ($sectionDiffs as $company => $data) {
				$changeColor = $data['diff'] > 0 ? 'green' : 'red';
				$changeSymbol = $data['diff'] > 0 ? '+' : '';

				$sectionTable->addRow([
					$company,
					"<fg={$changeColor}>{$changeSymbol}" . round($data['diff'] * 100, 4) . '%',
					round($data['old'] * 100, 4) . '%',
					round($data['new'] * 100, 4) . '%'
				]);
			}

			$sectionTable->render();
			$output->writeln(''); // Add spacing between sections
		}
	}

	/**
	 * Convert section key to readable display name
	 */
	private function getSectionDisplayName(string $sectionKey): string {
		$displayNames = [
			'equity_and_equity_related' => 'Equity & Equity Related',
			'arbitrage' => 'Arbitrage',
			'b_reits' => 'REITs',
			'equity_foreign_investments' => 'Foreign Equity Investments',
			'certificate_of_deposit' => 'Certificate of Deposit',
			'commercial_paper' => 'Commercial Paper',
			'treasury_bill' => 'Treasury Bills',
			'mutual_fund_units' => 'Mutual Fund Units',
			'reverse_repo_treps' => 'Reverse Repo / TREPS',
		];

		return $displayNames[$sectionKey] ?? ucwords(str_replace('_', ' ', $sectionKey));
	}
}

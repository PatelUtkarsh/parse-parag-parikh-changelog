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

		// Compare key and value = find diff of % and list.
		$column_diff = [];
		$all_companies = array_unique(array_merge(array_keys($old_sheet), array_keys($new_sheet)));

		foreach ( $all_companies as $name ) {
			$old_percentage = $old_sheet[$name] ?? 0;
			$new_percentage = $new_sheet[$name] ?? 0;
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

		return Command::SUCCESS;
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

			// Check first 10 rows to find column headers
			for ($row = 1; $row <= 10; $row++) {
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
							strpos($lowerValue, 'security') !== false) {
							$nameColumn = $col;
						}
						if (strpos($lowerValue, '%') !== false ||
							strpos($lowerValue, 'percent') !== false ||
							strpos($lowerValue, 'weight') !== false) {
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

			$output = [];
			foreach ( $worksheet->getRowIterator() as $row ) {
				$cellIterator = $row->getCellIterator();
				$cellIterator->setIterateOnlyExistingCells( false );
				$cells = [];

				foreach ( $cellIterator as $cell ) {
					$value = $cell->getValue();
					$column = $cell->getColumn();

					if ($column === $nameColumn) {
						$cells['name'] = $value;
					} elseif ($column === $percentColumn) {
						$cells['percent'] = $value;
					}
				}

				if (!empty($cells) && isset($cells['name']) && isset($cells['percent'])) {
					$output[] = $cells;
				}
			}

			// Clear all array entries where percent is not numeric or name is empty
			$output = array_filter( $output, function ( $item ) {
				return !empty( $item['name'] ) &&
					   is_string($item['name']) &&
					   strlen(trim($item['name'])) > 2 &&
					   isset( $item['percent'] ) &&
					   is_numeric($item['percent']) &&
					   $item['percent'] > 0;
			} );

			// Convert percent to float if it's not already
			$output = array_map(function($item) {
				$item['percent'] = (float) $item['percent'];
				$item['name'] = trim($item['name']);
				return $item;
			}, $output);

			// Make key value array.
			$names = array_column($output, 'name');
			$percents = array_column($output, 'percent');

			if (count($names) !== count($percents) || empty($names)) {
				if ($try < 1 && is_array(self::SHEET_MAP[$sheetCode])) {
					$this->output->isVerbose() && $this->output->writeln('Unable to find valid data in sheet: ' . $sheetName . ' retrying with next sheet name.');
					return $this->get_excel_column_map($filename, $sheetCode, $try + 1);
				}
				return [];
			}

			$output = array_combine($names, $percents);

			// Remove items which contains any of following strings.
			$remove_strings = [
				'Total',
				'Last 3 year',
				'Last 1 year',
				'Last 5 year',
				'Since Inception',
				'Clearing Corporation of India Ltd',
				'Grand Total',
				'Sub Total',
				'TOTAL',
				'Equity & Equity related',
				'Listed / awaiting listing',
				'awaiting listing',
				'(MD ',
				'Tbill',
				'Days',
				'#'
			];

			foreach ( $remove_strings as $remove_string ) {
				$output = array_filter( $output, function ( $key ) use ( $remove_string ) {
					return stripos( $key, $remove_string ) === false;
				}, ARRAY_FILTER_USE_KEY );
			}

			$filtered_output = array_filter( $output );
			if ( empty( $filtered_output ) && $try < 1 && is_array( self::SHEET_MAP[ $sheetCode ] ) ) {
				$this->output->isVerbose() && $this->output->writeln( 'Unable to find any data in sheet: ' . $sheetName . ' retrying with next sheet name.' );
				return $this->get_excel_column_map( $filename, $sheetCode, $try + 1 );
			}

			return $filtered_output;

		} catch (\Exception $e) {
			$this->output->writeln('<error>Error processing Excel file: ' . $e->getMessage() . '</error>');
			return [];
		}
	}

	public function download_spreadsheet_in_temp_dir( string $url ): string|false {
		$userAgents = [
			'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
			'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
			'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
			'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:56.0) Gecko/20100101 Firefox/56.0',
			'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:56.0) Gecko/20100101 Firefox/56.0',
			'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_5) AppleWebKit/537.36 (KHTML, like Gecko) Safari/537.36',
			'Mozilla/5.0 (Windows NT 10.0; Win64; x64; Trident/7.0; AS; rv:11.0) like Gecko',
			'Mozilla/5.0 (Windows NT 6.1; Win64; x64; Trident/7.0; AS; rv:11.0) like Gecko',
			'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_5) AppleWebKit/603.2.4 (KHTML, like Gecko) Version/10.1.1 Safari/603.2.4',
			'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/15.15063',
		];

		$temp_dir  = sys_get_temp_dir();
		$filename  = basename( $url );
		$temp_file = $temp_dir . DIRECTORY_SEPARATOR . strtok( $filename, '?' );
		if ( file_exists( $temp_file ) ) {
			$this->output->isVerbose() && $this->output->writeln( 'File already exists: ' . $temp_file );
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
			fclose( $fp );
			if (file_exists($temp_file)) {
				unlink( $temp_file );
			}
			$this->output->isVerbose() && $this->output->writeln( 'Unable to download file: ' . $url . ' Error: ' . $e->getMessage() );
			return false;
		}

		fclose( $fp );

		if ( $http_code !== 200 ) {
			if (file_exists($temp_file)) {
				unlink( $temp_file );
			}
			$this->output->isVerbose() && $this->output->writeln( 'Unable to download file: ' . $url . ' HTTP Code: ' . $http_code );
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
			$current_date->modify( '-' . $month_diff . ' month' );
		}
		$last_day_of_month = $this->get_last_date_of_datetime( $current_date );
		$year              = $last_day_of_month->format( 'Y' );
		$month             = $last_day_of_month->format( 'F' );
		$day               = $last_day_of_month->format( 'd' );

		return sprintf( 'https://amc.ppfas.com/downloads/portfolio-disclosure/%1$s/PPFAS_Monthly_Portfolio_Report_%2$s_%3$s_%1$s.%4$s', $year, $month, $day, $extension );
	}
}

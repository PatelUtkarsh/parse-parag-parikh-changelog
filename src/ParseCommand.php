<?php

namespace Spock\PhpParseMutualFund;

use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Helper\Table;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;


class ParseCommand extends Command {

	private OutputInterface $output;

	const SHEET_MAP = [
		'flexi'  => 'PPFCF',
		'liquid' => 'PPFLP',
		'hybrid' => 'PPCHF',
		'tax'    => 'PPTSF',
	];

	public function configure() {
		$this->setName( 'parse' )
		     ->setDescription( 'Find diff of mutual fund stock holdings.' )
		     ->setHelp( 'This command help you generate diff of share changes per month for a PPFAS mutual fund' )
		     ->addArgument( 'month-diff-x', InputArgument::OPTIONAL, 'For comparing old sheet as X, As if Y - X.', 2 )
		     ->addOption( 'fund-name', 'f', InputArgument::OPTIONAL, 'Pick from ' . join( ',', array_keys( self::SHEET_MAP ) ), 'tax' )
		     ->addOption( 'month-diff-y', 'y', InputArgument::OPTIONAL, 'For comparing new sheet as Y, As if Y - X.', 1 );
	}

	public function execute( InputInterface $input, OutputInterface $output ) {
		$option    = $input->getOption( 'fund-name' );
		$month_old = intval( $input->getArgument( 'month-diff-x' ) );
		$month_new = intval( $input->getOption( 'month-diff-y' ) );
		$this->output = $output;

		$old_path = $this->download_handler( $month_old );
		$new_path = $this->download_handler( $month_new );

		if ( ! $old_path || ! $new_path ) {
			$output->writeln( 'Unable to download file.' );

			return Command::FAILURE;
		}

		$old_sheet = $this->get_excel_column_map( $old_path, $option );
		$new_sheet = $this->get_excel_column_map( $new_path, $option );

		// Compare key and value = find diff of % and list.
		$column_diff = [];
		foreach ( $new_sheet as $name => $percentage ) {
			if ( isset( $old_sheet[ $name ] ) ) {
				$column_diff[ $name ] = $percentage - $old_sheet[ $name ];
			} else {
				$column_diff[ $name ] = $percentage;
			}
		}
		// Sort column by value and print in table format.
		arsort( $column_diff );
		$table = new Table( $output );
		$table->setHeaders( [ 'Name', 'Percentage', 'Previous Holding', 'Current Holding' ] );
		foreach ( $column_diff as $name => $percentage ) {
			$table->addRow( [
				$name,
				round( $percentage * 100, 4 ) . '%',
				isset( $old_sheet[ $name ] ) ? $old_sheet[ $name ] * 100 : '',
				$new_sheet[ $name ] * 100
			] );
		}
		$table->render();

		return Command::SUCCESS;
	}

	public function download_handler( int $month ): string|bool {
		$url      = $this->generate_url( $month, 'xlsx' );
		$this->output->isDebug() && $this->output->writeln( 'Downloading file: ' . $url );
		$new_path = $this->download_spreadsheet_in_temp_dir( $url );
		if ( ! $new_path ) {
			$this->output->isDebug() && $this->output->writeln( 'Unable to download xlsx file, retrying with xls extension: ' . $url );
			// retry download with xlsx extension.
			$new_url = $this->generate_url( $month );

			return $this->download_spreadsheet_in_temp_dir( $new_url );
		}

		return $new_path;
	}

	public function get_excel_column_map( string $filename, string $sheet_name ) {
		// Extract file name extension.
		$inputFileType = \PhpOffice\PhpSpreadsheet\IOFactory::identify( $filename );
		$filterSubset  = new CustomFilterRowColumns();
		// Create a new Reader of the type that has been identified
		$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader( $inputFileType );
		$reader->setReadDataOnly( true );
		$reader->setLoadSheetsOnly( [ self::SHEET_MAP[ $sheet_name ] ] );
		$reader->setReadFilter( $filterSubset );
		$spreadsheet = $reader->load( $filename );
		// get sheet names.
		// $sheet_names = $spreadsheet->getSheetNames();
		// var_dump($sheet_names);
		$worksheet = $spreadsheet->getActiveSheet();

		foreach ( $worksheet->getRowIterator() as $row ) {
			$cellIterator = $row->getCellIterator();
			$cellIterator->setIterateOnlyExistingCells( false );
			$cells = [];
			// We only care about column A and G.
			// A is the name of the stock.
			// G is %.
			// We can ignore the rest.
			foreach ( $cellIterator as $cell ) {
				$value = $cell->getValue();
				switch ( $cell->getColumn() ) {
					case 'B':
						$cells['name'] = $value;
						break;
					case 'G':
						$cells['percent'] = $value;
						break;
				}
			}
			$output[] = $cells;
		}

		// Clear all array entries where percent is not float.
		$output = array_filter( $output, function ( $item ) {
			return is_float( $item['percent'] ) || empty( $item['name'] );
		} );

		// Make key value array.
		$output = array_combine( array_column( $output, 'name' ), array_column( $output, 'percent' ) );

		// Remove items which contains any of following strings.
		$remove_strings = [
			'Total',
			'Last 3 year',
			'Last 1 year',
		];
		foreach ( $remove_strings as $remove_string ) {
			$output = array_filter( $output, function ( $item ) use ( $remove_string ) {
				return strpos( $item, $remove_string ) === false;
			}, ARRAY_FILTER_USE_KEY );
		}

		return array_filter( $output );
	}

	public function download_spreadsheet_in_temp_dir( string $url ) {
		$temp_dir  = sys_get_temp_dir();
		$filename  = basename( $url );
		$temp_file = $temp_dir . DIRECTORY_SEPARATOR . strtok( $filename, '?' );
		if ( file_exists( $temp_file ) ) {
			return $temp_file;
		}
		$fp = fopen( $temp_file, 'w+' );
		$ch = curl_init( $url );
		// if 404 then do not download.
		curl_setopt( $ch, CURLOPT_NOBODY, true );
		curl_setopt( $ch, CURLOPT_RETURNTRANSFER, true );
		curl_setopt( $ch, CURLOPT_FILE, $fp );
		curl_setopt( $ch, CURLOPT_FOLLOWLOCATION, true );
		curl_exec( $ch );
		$http_code = curl_getinfo( $ch, CURLINFO_HTTP_CODE );
		curl_close( $ch );
		fclose( $fp );
		if ( $http_code !== 200 ) {
			unlink( $temp_file );

			return false;
		}

		return $temp_file;
	}

	public function get_last_date_of_datetime( \DateTime $current_date ) {
		return new \DateTime( date( "Y-m-t", $current_date->format( 'U' ) ) );
	}

	public function generate_url( int $month_diff, $extension = 'xls' ) {
		// Last month date object.
		$current_date = new \DateTime();
		if ( $month_diff !== 0 ) {
			$current_date->modify( '-' . $month_diff . ' month' );
		}
		$last_day_of_month = $this->get_last_date_of_datetime( $current_date );
		$year              = $last_day_of_month->format( 'Y' );
		$month             = $last_day_of_month->format( 'F' );
		$day               = $last_day_of_month->format( 'd' );

		$extension = 'xls';
		// If August 31, 2023, or later than use new url.
		if ( $last_day_of_month->format( 'U' ) >= strtotime( '2023-08-31' ) ) {
			$extension = 'xlsx';
		}

		return sprintf( 'https://amc.ppfas.com/downloads/portfolio-disclosure/%1$s/PPFAS_Monthly_Portfolio_Report_%2$s_%3$s_%1$s.%4$s', $year, $month, $day, $extension );
	}
}

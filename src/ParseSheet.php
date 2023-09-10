<?php

namespace Spock\PhpParseMutualFund;

use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Helper\Table;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;


class ParseSheet extends Command {

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
		     ->addArgument( 'oldurl', InputArgument::REQUIRED, 'Spreadsheet of previous month' )
		     ->addArgument( 'newurl', InputArgument::REQUIRED, 'Spreadsheet of this month' )
		     ->addOption( 'fund-name', 'f', InputArgument::OPTIONAL, 'Pick from ' . join( ',', array_keys( self::SHEET_MAP ) ), 'tax' );
	}

	public function execute( InputInterface $input, OutputInterface $output ) {
		// Get option value.
		$option  = $input->getOption( 'fund-name' );
		$oldpath = $this->download_spreadsheet_in_temp_dir( $input->getArgument( 'oldurl' ) );
		$output->writeln( 'File downloaded to ' . $oldpath );
		$newpath = $this->download_spreadsheet_in_temp_dir( $input->getArgument( 'newurl' ) );
		$output->writeln( 'File downloaded to ' . $newpath );

		$old_sheet = $this->get_excel_column_map( $oldpath, $option );
		$new_sheet = $this->get_excel_column_map( $newpath, $option );

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

	public function get_excel_column_map( string $filename, string $sheet_name ) {
		// Extract file name extension.
		$inputFileType = \PhpOffice\PhpSpreadsheet\IOFactory::identify( $filename );
		$filterSubset  = new MyReadFilter();
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
		curl_setopt( $ch, CURLOPT_FILE, $fp );
		curl_setopt( $ch, CURLOPT_FOLLOWLOCATION, true );
		curl_exec( $ch );
		curl_close( $ch );
		fclose( $fp );

		return $temp_file;
	}
}

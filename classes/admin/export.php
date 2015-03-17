<?php

class GF_EXCEL_Export {

	private static $objPHPExcel = null;

	public function __construct() {
		add_action( 'init', array( __CLASS__, 'init' ), 1 );
	}

	/*
	 * 
	 * Hook on init before GravityForms plugin to catch event.
	 * Launch our download action
	 * 
	 * @return void
	 * @author Alexandre Sadowski
	 */

	public static function init() {
		if ( is_admin() && class_exists( 'GFCommon' ) ) {
			if ( GFCommon::current_user_can_any( GFCommon::all_caps() ) ) {
				require_once(GFCommon::get_base_path() . "/export.php");
				self::maybe_export();
			}
		}
	}

	/*
	 * 
	 * Check if export_gf_excel is set on POST and launch Excel Export
	 * 
	 * @return void
	 * @author Alexandre Sadowski
	 */

	public static function maybe_export() {
		if ( isset( $_POST[ 'export_lead' ] ) && ( isset( $_POST[ 'export_gf_excel' ] ) && (int) $_POST[ 'export_gf_excel' ] === 1 ) ) {
			check_admin_referer( 'rg_start_export', 'rg_start_export_nonce' );
			//see if any fields chosen
			if ( empty( $_POST[ 'export_field' ] ) ) {
				GFCommon::add_error_message( __( 'Please select the fields to be exported', 'gv-excel' ) );
				return;
			}

			$form = RGFormsModel::get_form_meta( $_POST[ 'export_form' ] );

			/** Include PHPExcel */
			require_once GF_EXCEL_DIR . '/libraries/PHPExcel.php';

			// Create new PHPExcel object
			self::$objPHPExcel = new PHPExcel();

			// Set active sheet index to the first sheet, so Excel opens this as the first sheet
			self::$objPHPExcel->setActiveSheetIndex( 0 );
			$result = self::get_array_for_excel( $form );

			self::$objPHPExcel->getActiveSheet()->fromArray( $result, NULL, 'A1' );
			$filename = sanitize_title_with_dashes( $form[ 'title' ] ) . '-' . gmdate( 'Y-m-d', GFCommon::get_local_timestamp( time() ) ) . '.xlsx';

			// Redirect output to a client’s web browser (Excel2007)
			header( 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
			header( "Content-Disposition: attachment; filename=$filename" );
			header( 'Cache-Control: max-age=0' );
			// If you're serving to IE 9, then the following may be needed
			header( 'Cache-Control: max-age=1' );

			// If you're serving to IE over SSL, then the following may be needed
			header( 'Expires: Mon, 26 Jul 1997 05:00:00 GMT' ); // Date in the past
			header( 'Last-Modified: ' . gmdate( 'D, d M Y H:i:s' ) . ' GMT' ); // always modified
			header( 'Cache-Control: cache, must-revalidate' ); // HTTP/1.1
			header( 'Pragma: public' ); // HTTP/1.0

			$objWriter = PHPExcel_IOFactory::createWriter( self::$objPHPExcel, 'Excel2007' );
			$objWriter->save( 'php://output' );
			exit;
		}
	}

	/*
	 * 
	 * Create full array with results to generate Excel
	 * 
	 * @param $form
	 * @return (array) List of results
	 * 			array(
	 * 				array( '', '' ) //Line1
	 * 				array( '', '' ) //Line2
	 * 			);
	 * @author Alexandre Sadowski
	 */

	private static function get_array_for_excel( $form ) {
		$form_id = $form[ 'id' ];
		$fields	 = $_POST[ 'export_field' ];

		$start_date	 = empty( $_POST[ 'export_date_start' ] ) ? '' : GFExport::get_gmt_date( $_POST[ 'export_date_start' ] . ' 00:00:00' );
		$end_date	 = empty( $_POST[ 'export_date_end' ] ) ? '' : GFExport::get_gmt_date( $_POST[ 'export_date_end' ] . ' 23:59:59' );

		$search_criteria[ 'status' ]		 = 'active';
		$search_criteria[ 'field_filters' ]	 = GFCommon::get_field_filters_from_post();
		if ( !empty( $start_date ) ) {
			$search_criteria[ 'start_date' ] = $start_date;
		}

		if ( !empty( $end_date ) ) {
			$search_criteria[ 'end_date' ] = $end_date;
		}

		$sorting = array( 'key' => 'date_created', 'direction' => 'DESC', 'type' => 'info' );
		$form	 = GFExport::add_default_export_fields( $form );

		$entry_count = GFAPI::count_entries( $form_id, $search_criteria );

		$page_size	 = 100;
		$offset		 = 0;

		//writing header
		$final_array = array();
		$header		 = array();
		foreach ( $fields as $field_id ) {
			$field	 = RGFormsModel::get_field( $form, $field_id );
			$value	 = str_replace( '"', '""', GFCommon::get_label( $field, $field_id ) );

			$header[] = $value;
		}

		$final_array[] = $header;
		//paging through results for memory issues
		while ( $entry_count > 0 ) {
			$paging	 = array(
				'offset'	 => $offset,
				'page_size'	 => $page_size
			);
			$leads	 = GFAPI::get_entries( $form_id, $search_criteria, $sorting, $paging );
			$leads	 = apply_filters( 'gform_leads_before_export_$form_id', apply_filters( 'gform_leads_before_export', $leads, $form, $paging ), $form, $paging );
			foreach ( $leads as $lead ) {
				$line = array();
				foreach ( $fields as $field_id ) {
					switch ( $field_id ) {
						case 'date_created' :
							$lead_gmt_time	 = mysql2date( 'G', $lead[ 'date_created' ] );
							$lead_local_time = GFCommon::get_local_timestamp( $lead_gmt_time );
							$value			 = date_i18n( 'Y-m-d H:i:s', $lead_local_time, true );
							break;
						default :
							$long_text		 = '';
							if ( strlen( rgar( $lead, $field_id ) ) >= (GFORMS_MAX_FIELD_LENGTH - 10) ) {
								$long_text = RGFormsModel::get_field_value_long( $lead, $field_id, $form );
							}

							$value = !empty( $long_text ) ? $long_text : rgar( $lead, $field_id );

							$field		 = RGFormsModel::get_field( $form, $field_id );
							$input_type	 = RGFormsModel::get_input_type( $field );

							if ( $input_type == 'checkbox' ) {
								$value	 = GFFormsModel::is_checkbox_checked( $field_id, $headers[ $field_id ], $lead, $form );
								if ( $value === false )
									$value	 = '';
							}
							elseif ( $input_type == 'fileupload' && rgar( $field, 'multipleFiles' ) ) {
								$value = !empty( $value ) ? implode( ' , ', json_decode( $value, true ) ) : '';
							}

							$value = apply_filters( 'gform_export_field_value', $value, $form_id, $field_id, $lead );

							break;
					}
					$line[] = $value;
				}
				$final_array[] = $line;
			}

			$offset += $page_size;
			$entry_count -= $page_size;
		}
		return $final_array;
	}

}

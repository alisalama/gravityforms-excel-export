<?php
class GF_EXCEL_Admin_Main {

	public function __construct() {
		add_action( 'admin_init', array( __CLASS__, 'register_assets' ) );
		add_action( 'admin_head', array( __CLASS__, 'enqueue_assets' ) );
	}

	/*
	 * Register list of assets used by the plugin
	 * 
	 * @return void
	 * @author Alexandre Sadowski
	 */
	public static function register_assets() {
		wp_register_script( 'gf-excel-admin', GF_EXCEL_URL . 'assets/js/admin-gf-excel.js', array( 'jquery' ), GF_EXCEL_VERSION, true );
		wp_localize_script( 'gf-excel-admin', 'gf_excel_vars', array( 'checkbox_excel' => __( 'Export to Excel format (xlsx)', 'gv-excel' ) ) );
	}
	
	/*
	 * Load registered assets used on good page template
	 * 
	 * @return void
	 * @author Alexandre Sadowski
	 */
	public static function enqueue_assets() {
		if ( isset( $_GET[ 'page' ] ) && $_GET[ 'page' ] == 'gf_export' ) {
			wp_enqueue_script( 'gf-excel-admin' );
		}
	}

}

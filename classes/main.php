<?php
class GF_EXCEL_Main {

	public function __construct() {
		add_action( 'init', array( __CLASS__, 'init' ) );
	}
	
	/*
	 * Load translations
	 * 
	 * @return void
	 * @author Alexandre Sadowski
	 */
	public static function init() {
		load_plugin_textdomain( 'gv-excel', false, basename(GF_EXCEL_DIR). '/languages/' );
	}
	
	public static function activate(){
		
	}
	
	public static function deactivate(){
		
	}
}
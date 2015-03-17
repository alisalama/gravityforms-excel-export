<?php
/*
 Plugin Name:Gravity Forms Excel Export
 Version: 1.0
 Description: Add the ability to export your messages in Excel format (xls)
 Author: BeAPI
 Author URI: http://www.beapi.fr
 Text Domain: gv-excel
 Domain Path: /languages/
 Depends: Gravity Forms

  --------------
  Copyright 2014 - BeAPI Team (technique@beapi.fr)

This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 2 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
 */

// don't load directly
if ( !defined('ABSPATH') )
	die('-1');

// Plugin constants
define( 'GF_EXCEL_VERSION', '1.0' );

// Plugin URL and PATH
define( 'GF_EXCEL_URL', plugin_dir_url ( __FILE__ ) );
define( 'GF_EXCEL_DIR', plugin_dir_path( __FILE__ ) );

// Function for easy load files
function _gf_excel_load_files( $dir, $files, $prefix = '' ) {
	foreach ( $files as $file ) {
		if ( is_file( $dir . $prefix . $file . ".php" ) ) {
			require_once( $dir . $prefix . $file . ".php" );
		}
	}	
}

// Plugin admin classes
if ( is_admin() ) {
	_gf_excel_load_files( GF_EXCEL_DIR . 'classes/admin/', array( 'export', 'main' ) );
}

// Plugin client classes
_gf_excel_load_files( GF_EXCEL_DIR . 'classes/', array( 'main') );

// Plugin activate/desactive hooks
register_activation_hook( __FILE__, array( 'GF_EXCEL_Plugin', 'activate' ) );
register_deactivation_hook( __FILE__, array( 'GF_EXCEL_Plugin', 'deactivate' ) );

add_action( 'plugins_loaded', 'init_gf_excel_plugin' );
function init_gf_excel_plugin() {

	new GF_EXCEL_Main();
	// Admin
	if (is_admin()) {
		new GF_EXCEL_Admin_Main();
		new GF_EXCEL_Export();
	}
}
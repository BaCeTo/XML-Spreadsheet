<?php defined('IS_ADMIN_FLAG') OR ("No direct access allowed.");

/**
 *	Class for generating XML spreadsheet documents
 *	in Microsoft Excel format
 *
 *	@author Vasil Vasilev
 **/
class Excel_XML_builder extends stdClass {

	/**	 Constants	 **/
	const FILE_EXTENSION = 'xml';
	const NEW_LINE = "\n";
	const TAB_CHAR = "\t";
	const NOT_AVAILABLE = 'N/A';
	const DOUBLE_TAB = "\t\t";

	/**
	 *		Element Attributes
	 *
	 *	NOTE: The ss, and x attribute types mean they are part of that xml
	 *	namespace (i.e. x:tag, ss:another_tag)
	 **/
	protected $table_attr_ss = array('DefaultRowHeight', 'LeftCell',
									 'StyleID', 'TopCell');

	protected $table_attr_x = array('FullColumns', 'FullRows');

	protected $cell_types = array('Number', 'DateTime',
								  'Boolean', 'String', 'Error');

	/**
	 * 	Variable holds a filename to be used in the download function
	 *
	 *	@access private
	 *	@var string
	 **/
	private $_filename = '';

	/**
	 *	Variable holds generated xml content
	 *
	 *	@access private
	 *	@var string
	 **/
	private $_contents = '';


	/**
	 *	Styles holder
	 *	@access private
	 *	@var array
	 **/
	private $_styles = array();

	/**
	 *	Variable to get the DOM structure of the document
	 *
	 *	@access private
	 *	@var array
	 **/
	private $_xml_holder = array();


	/**
	 *	This var will hold references to the current worksheet
	 *
	 *	@access private
	 *	@var string/integer
	 **/
	private $_current_worksheet = array();

	/**
	 *	Variable holds reference to the current table
	 *
	 *	@access private
	 *	@var pointer
	 **/
	private $_table_ref = NULL;

	/**
	 *	Reference to the current row
	 **/
	private $_row_ref = NULL;


	/**
	 *	Column Width
	 *
	 *	@access private
	 *	@var integer
	 **/
    private $_column_width = 200;


	public function __construct() {

		$this->_current_worksheet = NULL;
		$this->_table_ref = NULL;
		$this->_row_ref = NULL;

		$this->_filename = '';
		$this->_contents = '';

		$this->_styles = array();
		$this->_xml_holder = array();
	}


	/**
	 *	Method to apply the headers to the generated XML
	 *
	 *	@access private
	 **/
	private function _headers() {
		$styles = '';

		if ( is_array( $this->_styles) &&  ! empty($this->_styles) )
			$styles = implode('', $this->_styles );

		$styles = self::NEW_LINE . $styles;

		$header = <<<EOF
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<ss:Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
    <ss:ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
        <WindowHeight>12525</WindowHeight>
        <WindowWidth>15195</WindowWidth>
        <WindowTopX>480</WindowTopX>
        <WindowTopY>120</WindowTopY>
        <ActiveSheet>0</ActiveSheet>
        <ProtectStructure>False</ProtectStructure>
        <ProtectWindows>False</ProtectWindows>
    </ss:ExcelWorkbook>
    <ss:Styles>
        <ss:Style ss:ID="Default" ss:Name="Normal">
            <ss:Alignment ss:Vertical="Bottom"/>
            <ss:Borders/>
			<ss:Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
			<ss:Interior/>
        </ss:Style>
        <ss:Style ss:ID="bold">
            <ss:Font ss:Bold="1" ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
        </ss:Style>{$styles}</ss:Styles>
EOF;
		return $header . self::NEW_LINE;
	}


	/**
	 *	Method to apply the footers to the generated XML
	 *
	 *	@access private
	 **/
	private function _footer() {
		$footer = '</ss:Workbook>';
		return $footer;
	}


	/**		API FUNCTIONS 	**/

    /**
    * 	Add style to the header
    *
    * @param string $style_id: id of the style the cells will reference to
    * @param array $parameters: array with parameters
    */
    public function add_style($style_id, $parameters) {

		$font_options = array();
		$fill_options = array();
		$border_options = array('left' => array(),
								'top' => array(),
								'right' => array(),
								'bottom' => array()
								);
		$alignment_options = array();

        foreach ($parameters as $index => $data) {

            switch ($index) {

				/**	Font Options	*/
                case 'size':
                    $font_options['ss:Size'] = $data;
                break;

                case 'font':
                    $font_options['ss:FontName'] = $data;
                break;

                case 'color':
                case 'colour':
                    $font_options['ss:Color'] = $data;
                break;

                case 'bold':
                    $font_options['ss:Bold'] = $data;
                break;

                case 'italic':
                    $font_options['ss:Italic'] = $data;
                break;

                case 'strike':
                    $font_options['ss:StrikeThrough'] = $data;
                break;

				/**	 Fill Options	*/
			    case 'bgcolor':
                    $fill_options['ss:Color'] = $data;
                break;

				/**		Borders parsing */

				// Border Colors
				case 'border-all-color':
				case 'border-all-colour':
					$border_options['left']['ss:Color'] = $data;
					$border_options['right']['ss:Color'] = $data;
					$border_options['top']['ss:Color'] = $data;
					$border_options['bottom']['ss:Color'] = $data;
				break;

				case 'border-left-color':
				case 'border-left-colour':
					$border_options['left']['ss:Color'] = $data;
				break;

				case 'border-right-color':
				case 'border-right-colour':
					$border_options['right']['ss:Color'] = $data;
				break;

				case 'border-top-color':
				case 'border-top-colour':
					$border_options['top']['ss:Color'] = $data;
				break;

				case 'border-bottom-color':
				case 'border-bottom-colour':
					$border_options['bottom']['ss:Color'] = $data;
				break;

				// Border weight
				case 'border-all-weight':
					$border_options['left']['ss:Weight'] = $data;
					$border_options['right']['ss:Weight'] = $data;
					$border_options['top']['ss:Weight'] = $data;
					$border_options['bottom']['ss:Weight'] = $data;
				break;

				case 'border-left-weight':
					$border_options['left']['ss:Weight'] = $data;
				break;

				case 'border-right-weight':
					$border_options['right']['ss:Weight'] = $data;
				break;

				case 'border-top-weight':
					$border_options['top']['ss:Weight'] = $data;
				break;

				case 'border-bottom-weight':
					$border_options['bottom']['ss:Weight'] = $data;
				break;

				// Border line-style
				case 'border-all-line-style':
					$border_options['left']['ss:LineStyle'] = $data;
					$border_options['right']['ss:LineStyle'] = $data;
					$border_options['top']['ss:LineStyle'] = $data;
					$border_options['bottom']['ss:LineStyle'] = $data;
				break;

				case 'border-left-line-style':
					$border_options['left']['ss:LineStyle'] = $data;
				break;

				case 'border-right-line-style':
					$border_options['right']['ss:LineStyle'] = $data;
				break;

				case 'border-top-line-style':
					$border_options['top']['ss:LineStyle'] = $data;
				break;

				case 'border-bottom-line-style':
					$border_options['bottom']['ss:LineStyle'] = $data;
				break;

				/**	Alignment options */
				case 'alignment-horizontal':
					$alignment_options['ss:Horizontal'] = $data;
				break;

				case 'alignment-reading-order':
					$alignment_options['ss:ReadingOrder'] = $data;
				break;

				case 'alignment-vertical':
					$alignment_options['ss:Vertical'] = $data;
				break;

				default:
				break;
            }
        }

		$interior = '';
		$font = '';
		$borders = '';
		$alignment = '';

	    if ( ! empty($fill_options) )  {
			$fill_temp = '';
			foreach ($fill_options as $attr => $value)
                $fill_temp .= ' ' . $attr . '="' . $value . '"';

			$interior = self::DOUBLE_TAB . self::TAB_CHAR
					.'<ss:Interior ss:Pattern="Solid"' . $fill_temp
					.' />' . self::NEW_LINE;
		}

        if ( ! empty($font_options) ) {
			$font_temp = '';
            foreach ($font_options as $attr => $value)
                $font_temp .= ' ' . $attr . '="' . $value . '"';

            $font = self::DOUBLE_TAB . self::TAB_CHAR
					. '<ss:Font'. $font_temp
					.' />' . self::NEW_LINE;
        }

		if ( ! empty($alignment_options) ) {
			$align_temp = '';
			foreach ($alignment_options as $attr => $value)
				$align_temp .= ' '. $attr .'="' . $value .'"';
			$alignment = self::DOUBLE_TAB . self::TAB_CHAR
				. '<ss:Alignment'. $align_temp
				.' />' . self::NEW_LINE;
			}

		if ( 	! empty($border_options['left']) || ! empty($border_options['right'])
			 || ! empty($border_options['top']) || ! empty($border_options['bottom']) ) {

			$borders =  self::DOUBLE_TAB . self::TAB_CHAR .'<ss:Borders>' . self::NEW_LINE;

			foreach ($border_options as $direction => $options) {
				if (! empty($options) ) {
					$border_temp = self::DOUBLE_TAB . self::DOUBLE_TAB
					. '<ss:Border ss:Position="'
					. ucfirst(strtolower($direction)) .'"';

					foreach ($options as $attr => $value)
						$border_temp .= ' '. $attr .'="' . $value .'"';
					$borders .= $border_temp . ' />' . self::NEW_LINE;
				}
			}
			$borders .=  self::DOUBLE_TAB . self::TAB_CHAR .'</ss:Borders>' . self::NEW_LINE;
		}

        $this->_styles[] = self::DOUBLE_TAB
			.'<ss:Style ss:ID="'. $style_id .'" ss:Name="'. ucfirst(strtolower($style_id)).'">'. self::NEW_LINE
            . $interior . $font . $alignment . $borders
			. self::DOUBLE_TAB
			.'</ss:Style>' . self::NEW_LINE;
    }

	/**
	 *	Method to set the column width
	 *
	 *	@access public
	 *	@var integer
	 **/
	public function set_column_width($width = 0) {
		$width = (int) $width;

		if (! ( $width > 0 ) )
			$width = (int) $this->_column_width;

		$this->_column_width =  $width;
	}


	/**
	 *	The method creates a new workspace element in the XML DOM
	 *
	 *	@access public
	 *	@param string
	 **/
	public function add_worksheet($worksheet_name = '') {

		$worksheet_name = trim($worksheet_name);

		if ( ! strlen($worksheet_name) )
			$worksheet_name = sizeof($this->_xml_holder) + 1;

		$this->_xml_holder[$worksheet_name] = array(
			'name' => $worksheet_name,
			'table' => array(),
			'options' => '' // some special options
		);

		$this->_current_worksheet = $worksheet_name;
	}


	/**
	 *	Method closes the internal link to the current workspace
	 *
	 *	@access public
	 **/
	public function close_worksheet() {
		$this->_current_worksheet = NULL;
	}


	/**
	 *	Method creates a new table
	 *
	 *	Note: The current schema allows only one table per worksheet
	 *
	 *	@access public
	 *	@param
	 **/
	public function create_table($options = array()) {

		if ($this->_table_ref)
			unset($this->_table_ref);

		//make sure we have a worksheet
		if ( is_null($this->_current_worksheet) )
			$this->add_worksheet();

		$this->_xml_holder[ $this->_current_worksheet ]['table'] = array(
			'options' => '',
			'contents' => array()	//special options
		);

		if ( is_array($options) && ! empty($options) ) {

			$internal_options = '';

			foreach ($options as $key => $value ) {

				if ( in_array($key,	$this->table_attr_ss) )
					$internal_options .= ' ss:'. $key . '="'. $value .'"';

				if ( in_array($key,	$this->table_attr_x) )
					$internal_options .= ' x:'. $key . '="'. $value .'"';
			}

			$this->_xml_holder[ $this->_current_worksheet ]['table']['options'] = $internal_options;
		}

		// set a reference, as it is easier to work with
		$this->_table_ref = & $this->_xml_holder[ $this->_current_worksheet ]['table'];
	}


	/**
	 *	Add row to the table
	 *
	 *	@access public
	 **/
	public function add_row() {

		if ($this->_row_ref)
			unset($this->_row_ref);

		// make sure we have a table
		if ( ! $this->_table_ref )
			$this->create_table();

		$this->_table_ref['contents'][] = array();

		// set a reference to the current row
		$this->_row_ref = & $this->_table_ref['contents'][ sizeof($this->_table_ref['contents'])-1 ];

	}


	/**
	 *	Add a cell to the table
	 *
	 *	@access public
	 *
	 *	@var string content
	 *	@var string type
	 *	@var integer merge default 1
	 *
	 **/
	public function add_cell($content, $type = 'String' , $style = '', $merge = 0, $leave_blank = 0) {

		$content = htmlspecialchars( trim($content) );
		$type = ucfirst( strtolower($type) );
		$leave_blank = !! $leave_blank;
		$style = trim($style);
		$merge = (int) $merge;


		if ( ! in_array($type, $this->cell_types) )
			$type = 'String'; // the default type

		if ( ! strlen($content) && ! $leave_blank)
			$content = self::NOT_AVAILABLE;

		$data = '<ss:Data ss:Type="'. $type .'">'. $content . '</ss:Data>';

		$cell_attr = '';
		if ( strlen( trim($style) ) )
			$cell_attr .= ' ss:StyleID="'. trim( $style ) .'"';

		if ( (int) $merge > 0)
			$cell_attr .= ' ss:MergeAcross="'. (int) $merge .'"';

		$cell = self::DOUBLE_TAB . self::TAB_CHAR .'<ss:Cell' . $cell_attr. '>'.$data.'</ss:Cell>' . self::NEW_LINE;

		// push the cell into the array
		$this->_row_ref[] = $cell;
	}


	/**
	 *	Set filename for the force download
	 *
	 *	@param string
	 **/
	public function set_filename( $filename = '') {

		if ( preg_match('#^[^\/:*?"<>|]{1,255}?$#i', $filename) ) {
			$_filename  = $filename;

			if (preg_match('#^[^\/:*?"<>|]+\.xml$#i',$_filename) )
				$_filename = substr($_filename, 0, -4);	//strip the extension

			$this->_filename = trim($_filename)  . '.' . self::FILE_EXTENSION;
		} else {
			trigger_error("Whoops! Please select an appropriate filename.", E_ERROR);
		}

	}



	/**
	 *	Function will force direct printing of the generated xml structure
	 *
	 *	@access public
	 **/
	public function get_xml() {

		$this->_contents = '';

		// add header
		$this->_contents .= $this->_headers();
		// iterate through
		foreach ( $this->_xml_holder as $worksheet) {

			//Create worksheet
			$this->_contents .= '<ss:Worksheet ss:Name="' . $worksheet['name'] . '">' . self::NEW_LINE;
			// Create table with options

			$table_contents = '';
			$longest_row = 0;
			$rowIndex = 0;

			foreach ($worksheet['table']['contents'] as $rowContents) {
				// Create rows

				//calculate the size of the longest row in the table
				if (sizeof($rowContents) > $longest_row)
					$longest_row = sizeof($rowContents);

				$rowIndex ++;

				$table_contents .= self::TAB_CHAR .self::TAB_CHAR .'<ss:Row>' . self::NEW_LINE;
				foreach ($rowContents as $eachCell) {
					//Attach Cells
					$table_contents .= $eachCell;
				}
				$table_contents .= self::TAB_CHAR .self::TAB_CHAR ."</ss:Row>" . self::NEW_LINE;
			}

			$this->_contents .= self::TAB_CHAR . '<ss:Table ss:ExpandedRowCount="'.
					(int) sizeof($worksheet['table']['contents']) . '"'
					.' ss:DefaultColumnWidth="'. (int) $this->_column_width .'"'
					.' ss:ExpandedColumnCount="'.	(int) $longest_row . '"'.
					$worksheet['table']['options'] . '>' . self::NEW_LINE;
			$this->_contents .= $table_contents;
			$this->_contents .= self::TAB_CHAR .'</ss:Table>' . self::NEW_LINE;

			$this->_contents .= '</ss:Worksheet>' . self::NEW_LINE;
		}

		$this->_contents .= $this->_footer();
		return $this->_contents;
	}


	/**
	 *	Function forces downloading of the generated xml structure
	 *
	 *	@access public
	 **/
	public function download() {

		$this->get_xml();

	    if ( strlen($this->_contents) ) {

			header("Cache-Control: public, must-revalidate");
			header("Pragma: no-cache");
			header("Content-Length: ". strlen($this->_contents) );
			header("Content-Type: application/vnd.ms-excel");
			header('Content-Disposition: attachment; filename="'. $this->_filename .'"');
			header("Content-Transfer-Encoding: binary");
			print $this->_contents;
		}
        exit;
    }

}

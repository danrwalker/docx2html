<?php

	/**
	 * Convert docx files into HTML
	 * @author Dan Walker
	 * @link http://danrwalker.co.uk
	 * @reference Used XML to Array conversion taken from TwistPHP Framework (https://twistphp.com)
	 */
	class docx2html{

		protected $dirDocxPath = null;
		protected $strRawXML = null;
		protected $arrRawData = null;

		public function __construct(){

		}

		/**
		 * Load the docx file into the system
		 * @param $dirPathToDocx
		 * @throws Exception
		 */
		public function load($dirPathToDocx){

			if(file_exists($dirPathToDocx)){
				$this->dirDocxPath = $dirPathToDocx;
				$this->expandDocx();
			}else{
				throw new Exception(sprintf("docx2html: Invalid file path or file does not exist: %s",$dirPathToDocx),1);
			}
		}

		/**
		 * Expand the docx file into a usable raw XML format
		 * @throws Exception
		 */
		protected function expandDocx(){

			$this->strRawXML = '';
			$resDocxZip = zip_open($this->dirDocxPath);

			if(is_resource($resDocxZip)){

				while($resDocxEntry = zip_read($resDocxZip)){

					//Check to see that the entry is readable
					if(zip_entry_open($resDocxZip, $resDocxEntry) && zip_entry_name($resDocxEntry) == "word/document.xml"){
						$this->strRawXML .= zip_entry_read($resDocxEntry, zip_entry_filesize($resDocxEntry));
					}

					zip_entry_close($resDocxEntry);
				}

				zip_close($resDocxZip);

				if($this->strRawXML == ''){
					throw new Exception(sprintf("docx2html: No readable data found in docx file: %s",$this->dirDocxPath),3);
				}

			}else{
				throw new Exception(sprintf("docx2html: Failed to read docx file or invalid format: %s",$this->dirDocxPath),2);
			}
		}

		/**
		 * Process the XML into an initial linear (raw) array
		 */
		protected function processXML(){

			//If the xml data is a string turn it into an array
			if(is_string($this->strRawXML)){

				//Parse the raw XML into a raw XML Array (needs to be processed further)
				$resXML = xml_parser_create();
				xml_parse_into_struct($resXML, $this->strRawXML, $this->arrRawData);
				xml_parser_free($resXML);

				return $this->expandRawArray();
			}
		}

		protected $arrLevelMarker = array();
		protected $intKeyPosition = 0;

		/**
		 * Expand the linear (raw) array into a usable multi-level array
		 * @param $intCurrentKey
		 * @return array
		 */
		protected function expandRawArray($intCurrentKey = -1){

			$arrXmlData = array();
			$intCloseLevel = null;
			$blEndLoop = false;

			foreach($this->arrRawData as $intKey => $arrEachElement){

				//Only start processing once reach current array location
				//When completing an 'open' a skip value will be set (current array key)
				//The previous foreach loop will then ignore all of the child elements that have just been processed
				if($intKey > $intCurrentKey){

					//Check to see if a close level has been set, is bigger than current level return to previous iteration
					if(!is_null($intCloseLevel) && $intCloseLevel >= $arrEachElement['level']){
						$blEndLoop = true;
						$intKey = $intKey -1;
					}else{

						//Null the close level as it is not relevant at this point
						$intCloseLevel = null;

						$arrAttributes = array();

						//Process the attributes, lower case all the keys
						if(array_key_exists('attributes',$arrEachElement)){
							foreach($arrEachElement['attributes'] as $strKey => $strValue){
								$arrAttributes[strtolower($strKey)] = $strValue;
							}
						}

						switch($arrEachElement['type']){

							/**
							 * Process opening Tags
							 */
							case'open':

								$arrOpenContent = array();

								//If their is no data then don't add any to the field 'content'
								if(array_key_exists('value',$arrEachElement) && trim($arrEachElement['value']) != ''){
									$arrOpenContent[] = $arrEachElement['value'];
								}

								//Jump into the opening tag until the corresponding closing tag is found
								$arrChildData = $this->expandRawArray($intKey);

								//Grab the child data and update the current process location
								$arrOpenContent = array_merge($arrOpenContent,$arrChildData['data']);
								$intCurrentKey = $arrChildData['skip'];

								$arrXmlData[] = array(
									'level' => $arrEachElement['level'],
									'tag' => strtolower($arrEachElement['tag']),
									'attributes' => $arrAttributes,
									'content' => $arrOpenContent,
								);

								break;

							/**
							 * Process closing tags
							 */
							case'close':

								$intCloseLevel = $arrEachElement['level'];
								break;

							/**
							 * Process content data
							 */
							case'cdata':

								if(array_key_exists('value',$arrEachElement) && trim($arrEachElement['value']) != ''){
									$arrXmlData[] = $arrEachElement['value'];
								}

								break;

							/**
							 * Process a complete tag as a whole
							 */
							case'complete':

								$arrOpenContent = array();

								//If their is no data then don't add any to the field 'content'
								if(array_key_exists('value',$arrEachElement) && trim($arrEachElement['value']) != ''){
									$arrOpenContent[] = $arrEachElement['value'];
								}

								$arrXmlData[] = array(
									'level' => $arrEachElement['level'],
									'tag' => strtolower($arrEachElement['tag']),
									'attributes' => $arrAttributes,
									'content' => $arrOpenContent,
								);

								break;
						}
					}
				}

				//Exit the loop, this is used when the corresponding 'close' tag is found
				if($blEndLoop == true){
					break;
				}
			}

			//Send back the data and the Skip Key
			return array('data' => $arrXmlData,'skip' => $intKey);
		}

		/**
		 * Export the XML data gathered from the docx file
		 * @return null
		 * @throws Exception
		 */
		public function xml(){

			if(is_null($this->dirDocxPath)){
				throw new Exception("docx2html: A docx file must be loaded before XML content can be output",4);
			}

			return $this->strRawXML;
		}

		/**
		 * Export the data as plain text without formatting
		 * @return string
		 * @throws Exception
		 */
		public function plain(){

			if(is_null($this->dirDocxPath)){
				throw new Exception("docx2html: A docx file must be loaded before Plain content can be output",5);
			}

			$strPlainContent = str_replace('</w:r></w:p></w:tc><w:tc>',' ',$this->strRawXML);
			$strPlainContent = str_replace('</w:r></w:p>',"\r\n",$strPlainContent);

			return strip_tags($strPlainContent);
		}

		/**
		 * Export the data as formatted HTML
		 * @return string
		 * @throws Exception
		 */
		public function html(){

			if(is_null($this->dirDocxPath)){
				throw new Exception("docx2html: A docx file must be loaded before HTML content can be output",6);
			}

			//@todo Need more work
			$arrData = $this->processXML();
			$strHTML = print_r($arrData,true);

			return $strHTML;
		}

	}
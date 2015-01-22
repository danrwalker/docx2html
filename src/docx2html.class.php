<?php

	/**
	 * Convert docx files into HTML
	 * @author Dan Walker
	 * @link http://danrwalker.co.uk
	 */
	class docx2html{

		protected $dirDocxPath = null;
		protected $mxdRawXML = null;

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

			$this->mxdRawXML = '';
			$resDocxZip = zip_open($this->dirDocxPath);

			if(is_resource($resDocxZip)){

				while($resDocxEntry = zip_read($resDocxZip)){

					//Check to see that the entry is readable
					if(zip_entry_open($resDocxZip, $resDocxEntry) && zip_entry_name($resDocxEntry) == "word/document.xml"){
						$this->mxdRawXML .= zip_entry_read($resDocxEntry, zip_entry_filesize($resDocxEntry));
					}

					zip_entry_close($resDocxEntry);
				}

				zip_close($resDocxZip);

				if($this->mxdRawXML == ''){
					throw new Exception(sprintf("docx2html: No readable data found in docx file: %s",$this->dirDocxPath),3);
				}

			}else{
				throw new Exception(sprintf("docx2html: Failed to read docx file or invalid format: %s",$this->dirDocxPath),2);
			}
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

			return $this->mxdRawXML;
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

			$strPlainContent = str_replace('</w:r></w:p></w:tc><w:tc>',' ',$this->mxdRawXML);
			$strPlainContent = str_replace('</w:r></w:p>',"\r\n",$strPlainContent);

			return strip_tags($strPlainContent);
		}

	}
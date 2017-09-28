<?php
/*
*  Module written by Herman Kuiper <herman@ozuzo.net>
*
*  License Information:
*
*    Spreadsheet_Excel_Writer:  A library for generating Excel Spreadsheets
*    Copyright (c) 2002-2003 Xavier Noguer xnoguer@rezebra.com
*
*    This library is free software; you can redistribute it and/or
*    modify it under the terms of the GNU Lesser General Public
*    License as published by the Free Software Foundation; either
*    version 2.1 of the License, or (at your option) any later version.
*
*    This library is distributed in the hope that it will be useful,
*    but WITHOUT ANY WARRANTY; without even the implied warranty of
*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
*    Lesser General Public License for more details.
*
*    You should have received a copy of the GNU Lesser General Public
*    License along with this library; if not, write to the Free Software
*    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

/**
* Baseclass for generating Excel DV records (validations)
*
* @author   Herman Kuiper
* @category FileFormats
* @package  Spreadsheet_Excel_Writer
*/
class Spreadsheet_Excel_Writer_Validator
{
    // Possible operator types
    CONST OP_BETWEEN =    0x00;
    CONST OP_NOTBETWEEN = 0x01;
    CONST OP_EQUAL =      0x02;
    CONST OP_NOTEQUAL =   0x03;
    CONST OP_GT =         0x04;
    CONST OP_LT =         0x05;
    CONST OP_GTE =        0x06;
    CONST OP_LTE =        0x07;

    public $_type;
    public $_style;
    public $_fixedList;
    public $_blank;
    public $_incell;
    public $_showprompt;
    public $_showerror;
    public $_title_prompt;
    public $_descr_prompt;
    public $_title_error;
    public $_descr_error;
    public $_operator;
    public $_formula1;
    public $_formula2;
    /**
    * The parser from the workbook. Used to parse validation formulas also
    * @var Spreadsheet_Excel_Writer_Parser
    */
    public $_parser;

    /**
     * Spreadsheet_Excel_Writer_Validator constructor.
     * @param $parser
     */
    public function __construct($parser)
    {
        $this->_parser       = $parser;
        $this->_type         = 0x01; // FIXME: add method for setting datatype
        $this->_style        = 0x00;
        $this->_fixedList    = false;
        $this->_blank        = false;
        $this->_incell       = false;
        $this->_showprompt   = false;
        $this->_showerror    = true;
        $this->_title_prompt = "\x00";
        $this->_descr_prompt = "\x00";
        $this->_title_error  = "\x00";
        $this->_descr_error  = "\x00";
        $this->_operator     = 0x00; // default is equal
        $this->_formula1     = '';
        $this->_formula2     = '';
    }

    /**
     * @param string $promptTitle
     * @param string $promptDescription
     * @param bool $showPrompt
     */
    public function setPrompt($promptTitle = "\x00", $promptDescription = "\x00", $showPrompt = true)
   {
      $this->_showprompt = $showPrompt;
      $this->_title_prompt = $promptTitle;
      $this->_descr_prompt = $promptDescription;
   }

    /**
     * @param string $errorTitle
     * @param string $errorDescription
     * @param bool $showError
     */
    public function setError($errorTitle = "\x00", $errorDescription = "\x00", $showError = true)
   {
      $this->_showerror = $showError;
      $this->_title_error = $errorTitle;
      $this->_descr_error = $errorDescription;
   }

    /**
     *
     */
    public function allowBlank()
   {
      $this->_blank = true;
   }

    /**
     *
     */
    public function onInvalidStop()
   {
      $this->_style = 0x00;
   }

    /**
     *
     */
    public function onInvalidWarn()
    {
        $this->_style = 0x01;
    }

    /**
     *
     */
    public function onInvalidInfo()
    {
        $this->_style = 0x02;
    }

    /**
     * @param $formula
     * @return bool
     * @throws \InvalidArgumentException
     */
    public function setFormula1($formula)
    {
        // Parse the formula using the parser in Parser.php
        $this->_parser->parse($formula);
        $this->_formula1 = $this->_parser->toReversePolish();

        return true;
    }

    /**
     * @param $formula
     * @return bool
     * @throws \InvalidArgumentException
     */
    public function setFormula2($formula)
    {
        // Parse the formula using the parser in Parser.php
        $this->_parser->parse($formula);
        $this->_formula2 = $this->_parser->toReversePolish();

        return true;
    }

    /**
     * @return int
     */
    protected function _getOptions()
    {
        $options = $this->_type;
        $options |= $this->_style << 3;
        if ($this->_fixedList) {
            $options |= 0x80;
        }
        if ($this->_blank) {
            $options |= 0x100;
        }
        if (!$this->_incell) {
            $options |= 0x200;
        }
        if ($this->_showprompt) {
            $options |= 0x40000;
        }
        if ($this->_showerror) {
            $options |= 0x80000;
        }
        $options |= $this->_operator << 20;

        return $options;
   }

    /**
     * @return string
     */
    protected function _getData()
   {
      $title_prompt_len = strlen($this->_title_prompt);
      $descr_prompt_len = strlen($this->_descr_prompt);
      $title_error_len = strlen($this->_title_error);
      $descr_error_len = strlen($this->_descr_error);

      $formula1_size = strlen($this->_formula1);
      $formula2_size = strlen($this->_formula2);

      $data  = pack("V", $this->_getOptions());
      $data .= pack("vC", $title_prompt_len, 0x00) . $this->_title_prompt;
      $data .= pack("vC", $title_error_len, 0x00) . $this->_title_error;
      $data .= pack("vC", $descr_prompt_len, 0x00) . $this->_descr_prompt;
      $data .= pack("vC", $descr_error_len, 0x00) . $this->_descr_error;

      $data .= pack("vv", $formula1_size, 0x0000) . $this->_formula1;
      $data .= pack("vv", $formula2_size, 0x0000) . $this->_formula2;

      return $data;
   }
}

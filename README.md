[![Build Status](https://travis-ci.org/dimarick/Spreadsheet_Excel_Writer.svg?branch=master)](https://travis-ci.org/pear/Spreadsheet_Excel_Writer)

This package is [Spreadsheet_Excel_Writer](http://pear.php.net/package/Spreadsheet_Excel_Writer) and has been migrated from [svn.php.net](https://svn.php.net/repository/pear/packages/Spreadsheet_Excel_Writer).

Please report all new issues [via the PEAR bug tracker](http://pear.php.net/bugs/search.php?cmd=display&package_name[]=Spreadsheet_Excel_Writer&order_by=ts1&direction=DESC&status=Open).

If this package is marked as unmaintained and you have fixes, please submit your pull requests and start discussion on the pear-qa mailing list.

To test, run

    $ phpunit

## Migration

- This package has no dependencies from any pear classes and uses php5 native exceptions for error handling.
Be careful, currently many unexpected issues can throw unexpected exceptions instead of silent fail

- Used standard psr-0 composer autoloader

- Constants OP_* moved to class Spreadsheet_Excel_Writer_Validator

- Constants SPREADSHEET_EXCEL_WRITER_* moved to class Spreadsheet_Excel_Writer_Parser

- Do not try to run it on php4


## Composer

This package comes with support for Composer.

To install from Composer

    $ composer require pear/spreadsheet_excel_writer

To install the latest development version

    $ composer require pear/spreadsheet_excel_writer:dev-master

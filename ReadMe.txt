How to "install" the Active Directory UDF						2020-03-26
--------------------------------------------------------------------------
This step is mandatory:
* Copy AD.au3 into the %ProgramFiles%\AutoIt3\Include, the AutoIt User Includes directory or the directory
  where your script is located

  THe following steps are optional:
For SciTE integration (user calltips and syntax highlighting)
* Run the User CallTip Manager from SciTE Config tool, tab Other Tools.
  For details please check: http://www.autoitscript.com/wiki/Adding_UDFs_to_AutoIt_and_SciTE

Help files and examples
* Copy the *.htm and the remaining *.au3 files to any directory you like. 
  You can't call the help and example scripts from the AutoIt help at the moment


How to use the Active Directory UDF								2011-03-07
--------------------------------------------------------------------------
* Every script has to have the following format:
  _AD_Open()					 ; open a onnection to the AD 
  calls to other _AD-functions	 ; query or manipulate the AD
  _AD_Close()					 ; close the connection to the AD
* Special characters:
  If you call a function with a Fully Qualified Domain Name (FQDN) as parameter you as a user have to make sure that 
  all special characters (#, /, comma etc.) are escaped with a preceding backslash. You can call function 
  _AD_FixSpecialChars(String,1) to escape or _AD_FixSpecialChars(string, 0) to unescape the special characters.
* SamAccountName:
  The SamAccountName of a computer is the computername with a trailing "$" e.g. @ComputerName & "$"


General															2020-03-26
--------------------------------------------------------------------------
* I noted some problems with "old" versions of the Date UDF (function _Date_Time_SystemTimeToDateTimeStr).
  So please use the most current version of AutoIt. I'm testing with 3.3.6.0 at the moment
* The UDF does not fully support the "granular password policy" feature implemented with Windows Server 2008.
  Invalid results for all password related functions might be returned. Please check the wiki for the implementation status.
* The UDF only supports one domain at a time (the domain to which you connect with _AD_Open). Cross domain scripts 
  are not possible with this UDF (at the moment)
* If you run Windows Vista or higher, UAC is enabled and you use functions that change the AD then you should insert #RequireAdmin into your script
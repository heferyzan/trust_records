Microsoft Word 2010/2013 TrustRecords Check
Inspiration: http://az4n6.blogspot.com/2016/02/more-on-trust-records-macros-and.html

This script checks for cases where users enable content in a Word 2010/2013 document file that was downloaded from the Internet or from an email. This script also prints the default Registry flag for the Macro Settings which indicates Macro flag is set for doc files.

Usage: python trustrecords.py -f NTUSER.dat

Requirements: python-registry from https://github.com/williballenthin/python-registry
Install python-registry: sudo python setup.py install

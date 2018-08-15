# Microsoft Word 2010/2013 TrustRecords Check
# Author: Heferyzan
# Inspiration: http://az4n6.blogspot.com/2016/02/more-on-trust-records-macros-and.html
#
# This script checks for cases where users enable content in a Word 2010/2013 document file
# that was downloaded from the Internet or from an email. This script also prints the
# default Registry flag for the Macro Settings which indicates Macro flag is set for doc files.
# 
# Usage: python trustrecords.py -f NTUSER.dat
#
# Requirements: python-registry from https://github.com/williballenthin/python-registry
# Install python-registry: sudo python setup.py install
#

"""
To Do:
[1]
Testing for Excel and other Microsoft Office files.

[2] 
Testing additional NTUSER.dat files. 
"""

import argparse
import sys
from Registry import Registry

def arg_parse():
	parser=argparse.ArgumentParser()
	parser.add_argument('-f', dest='file_name', help='type the filename', action='store', required=True)
	args = parser.parse_args()
	return args.file_name

def vbawarning_check(key): #this is currently screwed.. it needs to check for VBAWarnings
	print "[VBAWARNINGS CONFIG CHECK]"
	count = 0
	try:
		key_2013 = key.open("SOFTWARE\\Microsoft\\Office\\15.0\\Word\\Security")
		for subkey in key_2013.values():
			if subkey.name() == "VBAWarnings":
				print "VBAWarnings Config was Found. Current Setting: %s\n" % vbawarning_mapping(subkey.value())
				count += 1
	except Registry.RegistryKeyNotFoundException:
		pass
	try:
		key_2010 = key.open("SOFTWARE\\Microsoft\\Office\\14.0\\Word\\Security")
		for subkey in key_2010.values():
			if subkey.name() == "VBAWarnings":
				print "VBAWarnings Config for Word 2010 was Found. Current Setting: %s\n" % vbawarning_mapping(subkey.value())
				count += 1
	except Registry.RegistryKeyNotFoundException:
		pass
	if count == 0:
		print "No VBAWarnings Config Found. Default Setting is enabled: Disable all macros with notification.\n"

def vbawarning_mapping(dword):
	if dword == 1:
		return "Enable all macros."
	elif dword == 2:
		return "Disable all macros with notification."
	elif dword == 3:
		return "Disable all amcros except digitally signed macros."
	elif dword == 4:
		return "Disable all macros without notification."
	else:
		return "The DWORD value did not map to a text description of the flag."

def word2013(key):
	count = 0
	print "[WORD 2013 CONTENT ENABLED]"
	try:
		key_trust = key.open("SOFTWARE\\Microsoft\\Office\\15.0\\Word\\Security\\Trusted Documents\\TrustRecords")
	except Registry.RegistryKeyNotFoundException:
		print "Couldn't find Word 2013 data."
		sys.exit(-1)
	for subkey in key_trust.values():
		s = "".join(["%02x" % (ord(c)) for c in subkey.value()])
		if s[-8:] == "ffffff7f":
			print subkey.name() + "\n"
			count += 1
	if count == 0:
		print "Data found, but no records with content enabled.\n"

def word2010(key):
	count = 0
	print "[WORD 2010 CONTENT ENABLED]"
	try:
		key_trust = key.open("SOFTWARE\\Microsoft\\Office\\14.0\\Word\\Security\\Trusted Documents\\TrustRecords")
	except Registry.RegistryKeyNotFoundException:
		print "Couldn't find Word 2010 data."
		sys.exit(-1)
	for subkey in key_trust.values():
		s = "".join(["%02x" % (ord(c)) for c in subkey.value()])
		if s[-8:] == "ffffff7f":
			print subkey.name() + "\n"
			count += 1
	if count == 0:
		print "Data found, but no records with content enabled.\n"

def main(file_name):
	f = open(file_name, "rb")
	r = Registry.Registry(f)
	vbawarning_check(r)
	word2013(r)
	word2010(r)
	f.close()
	r.close()

if __name__ == "__main__":
	arguments = arg_parse()
	main(arguments)
# -*- coding: utf-8 -*-

#! python3

# USAGE
# python3 custom_invite_word_032718_1.py "GUESTNAMES.TXT"
# python3 custom_invite_word_032718_1.py "../tests/guests.txt"

import docx,sys

import logging
logging.basicConfig(level=logging.DEBUG, format=" %(asctime)s - %(levelname)s - %(message)s")
# logging.disable(logging.CRITICAL)

#####################################
# GUEST INFORMATION
#####################################

# gather guest information

guest_txt = sys.argv[1]

logging.debug( 'Guest list text file input is:  %s' % (guest_txt) )

def guest_list_maker(guest_txt):
	guest_list = []
	# open the guest list text file
	fileObj = open(guest_txt)
	guest_list = fileObj.readlines() # this should return a list
	logging.debug( 'The guest list extracted from text is:  ' )
	logging.debug( guest_list )
	# read each line and push the name into the list
	# return the guest_list to outside the function for use
	return guest_list

#####################################
# END GUEST INFORMATION
#####################################

#####################################
# WORD DOCUMENT
#####################################

# create new Word document

doc = docx.Document()

def construct_invite_1(guest_list):
	doc.add_paragraph('It would be a pleasure to have the company of')
	doc.add_paragraph("")


#####################################
# END WORD DOCUMENT
#####################################

#####################################
# EXECUTION
#####################################

guest_list_maker(guest_txt)

# save the final word doc

# doc.save('custom_invites_f.docx')

#####################################
# END EXECUTION
#####################################


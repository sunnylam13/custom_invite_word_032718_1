# -*- coding: utf-8 -*-

#! python3

# USAGE
# python3 custom_invite_word_032718_1.py "GUESTNAMES.TXT"
# python3 custom_invite_word_032718_1.py "../tests/guests.txt"

import docx

import logging
logging.basicConfig(level=logging.DEBUG, format=" %(asctime)s - %(levelname)s - %(message)s")
# logging.disable(logging.CRITICAL)

#####################################
# GUEST INFORMATION
#####################################

# gather guest information


#####################################
# END GUEST INFORMATION
#####################################

#####################################
# WORD DOCUMENT
#####################################

# create new Word document

doc = docx.Document()

def construct_invite_1():
	doc.add_paragraph('It would be a pleasure to have the company of')


#####################################
# END WORD DOCUMENT
#####################################

#####################################
# EXECUTION
#####################################

# save the final word doc

# doc.save('custom_invites_f.docx')

#####################################
# END EXECUTION
#####################################


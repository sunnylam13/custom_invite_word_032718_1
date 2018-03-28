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

def construct_invite_1(doc,guest_list):

	for name in guest_list:
		doc.add_paragraph('It would be a pleasure to have the company of')
		logging.debug(doc.paragraphs[0].text)
		doc.paragraphs[0].style = 'Subtitle'
		
		doc.add_paragraph(name)
		logging.debug(doc.paragraphs[1].text)
		doc.paragraphs[1].style = 'Title'
		
		document.add_paragraph().add_run('at')
		document.add_paragraph().add_run('11010 Memory Lane on the Evening of')
		# doc.add_paragraph('at 11010 Memory Lane on the Evening of')
		# logging.debug(doc.paragraphs[2].text)
		logging.debug(doc.paragraphs[2].runs[0].text)
		logging.debug(doc.paragraphs[2].runs[1].text)
		# logging.debug(doc.paragraphs[2].runs[0].text)
		# doc.paragraphs[2].runs[0].underline = True # 'at'
		# doc.paragraphs[2].runs[1].style['Heading3'] = True # '11010 Memory Lane on the Evening of'
		
		doc.add_paragraph('April 1st')
		logging.debug(doc.paragraphs[3].text)
		logging.debug(doc.paragraphs[3].runs[0].text)
		doc.paragraphs[3].style = 'Normal'
		
		# doc.add_paragraph("at 7 o'clock")
		document.add_paragraph().add_run('at')
		document.add_paragraph().add_run("7 o'clock")
		# logging.debug(doc.paragraphs[4].text)
		logging.debug(doc.paragraphs[4].runs[0].text)
		logging.debug(doc.paragraphs[4].runs[1].text)
		# doc.paragraphs[4].runs[0].underline.italic = True # 'at'
		# doc.paragraphs[4].runs[1].style = 'IntenseQuote' # '7 o'clock'
		
		doc.add_page_break() # add page break


#####################################
# END WORD DOCUMENT
#####################################

#####################################
# EXECUTION
#####################################

construct_invite_1( doc,guest_list_maker(guest_txt) )

# save the final word doc

doc.save('custom_invites_f.docx')

#####################################
# END EXECUTION
#####################################


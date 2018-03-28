# -*- coding: utf-8 -*-

#! python3

# USAGE
# python3 custom_invite_word_032718_1.py "GUESTNAMES.TXT"
# python3 custom_invite_word_032718_1.py "../tests/guests.txt"

import docx,sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt

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

# create styles

styles = doc.styles # whole doc's styles object

style_citation_1 = styles.add_style('Citation', WD_STYLE_TYPE.PARAGRAPH)
logging.debug( "The style name created is:  %s" % (style_citation_1.name) )
logging.debug( "The style type created is:  %s" % (style_citation_1.name) )
logging.debug( style_citation_1.type )

style_italic_emph_1 = styles.add_style('GuestItalics', WD_STYLE_TYPE.PARAGRAPH)
logging.debug( "The style name created is:  %s" % (style_italic_emph_1.name) )
logging.debug( "The style type created is:  %s" % (style_italic_emph_1.name) )
logging.debug( style_italic_emph_1.type )
font_style_italic_emph_1 = style_italic_emph_1.font
font_style_italic_emph_1.name = 'Calibri'
font_style_italic_emph_1.size = Pt(14)

def construct_invite_1(doc,guest_list):

	counter = 0

	for name in guest_list:
		doc.add_paragraph('It would be a pleasure to have the company of')
		logging.debug(doc.paragraphs[0].text)
		doc.paragraphs[0 + counter].style = 'Subtitle'
		logging.debug('Styles applied')
		
		# insert the guest's name
		doc.add_paragraph(name)
		logging.debug(doc.paragraphs[1].text)
		doc.paragraphs[1 + counter].style = 'Title'
		logging.debug('Styles applied')
		
		doc.add_paragraph().add_run('at')
		doc.paragraphs[2 + counter].add_run('11010 Memory Lane on the Evening of') # add to the same paragraph
		logging.debug(doc.paragraphs[2 + counter].runs[0].text)
		logging.debug(doc.paragraphs[2 + counter].runs[1].text)
		doc.paragraphs[2 + counter].runs[0].italic = True # 'at'
		# doc.paragraphs[2 + counter].runs[0].underline = WD_UNDERLINE.DOT_DASH # 'at'
		doc.paragraphs[2 + counter].runs[1].style = doc.styles['style_italic_emph_1'] # '11010 Memory Lane on the Evening of'
		logging.debug('Styles applied')

		doc.add_paragraph('April 1st')
		logging.debug(doc.paragraphs[3 + counter].text)
		logging.debug(doc.paragraphs[3 + counter].runs[0].text)
		doc.paragraphs[3 + counter].style = 'Normal'
		logging.debug('Styles applied')

		# doc.add_paragraph("at 7 o'clock")
		doc.add_paragraph().add_run('at')
		doc.paragraphs[4 + counter].add_run("7 o'clock") # add to the same paragraph
		logging.debug(doc.paragraphs[4 + counter].runs[0].text)
		logging.debug(doc.paragraphs[4 + counter].runs[1].text)
		doc.paragraphs[4 + counter].runs[0].underline.italic = True # 'at'
		doc.paragraphs[4 + counter].runs[1].style = doc.styles['style_italic_emph_1'] # '7 o'clock'
		logging.debug('Styles applied')

		doc.add_page_break() # add page break

		counter += 5 # increment by 5 positions so that you can style the next guest properly
		
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


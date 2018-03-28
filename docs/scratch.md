# Scratch Notes and Logs

## Tuesday, March 27, 2018 3:57 PM

you have a text file of guest names with one name per line...

	Prof. Plum
	Miss Scarlet
	Col. Mustard
	Al Sweigart
	Robocop

since Python-Docx can use only those styles that already exist in the Word doc, you'll have to first add these styles to a blank Word file and then open that file with Python-Docx...

should be one invite per page so page breaks will be needed...  after the last paragraph

that way you only need to open one Word document to print all invites at the same time

## Wednesday, March 28, 2018 9:50 AM

[Python File readlines() Method](https://www.tutorialspoint.com/python/file_readlines.htm)

	fileObject.readlines( sizehint );

Return Value
This method returns a list containing the lines.

Example
The following example shows the usage of readlines() method.

	This is 1st line
	This is 2nd line
	This is 3rd line
	This is 4th line
	This is 5th line

## Wednesday, March 28, 2018 10:07 AM

Decided to use Word default styles to save time rather than creating a Word doc with custom styles...

Too much hassle...

## Wednesday, March 28, 2018 10:12 AM

Ran into styling error...


	MacBook-Air:custom_invite_word sunnyair$ python3 custom_invite_word_032718_1.py "../tests/guests.txt"
	 2018-03-28 09:55:43,511 - DEBUG - Guest list text file input is:  ../tests/guests.txt
	 2018-03-28 09:55:43,542 - DEBUG - The guest list extracted from text is:
	 2018-03-28 09:55:43,542 - DEBUG - ['Prof. Plum\n', 'Miss Scarlet\n', 'Col. Mustard\n', 'Al Sweigart\n', 'Robocop']
	MacBook-Air:custom_invite_word sunnyair$ python3 custom_invite_word_032718_1.py "../tests/guests.txt"
	 2018-03-28 10:11:45,993 - DEBUG - Guest list text file input is:  ../tests/guests.txt
	 2018-03-28 10:11:46,031 - DEBUG - The guest list extracted from text is:
	 2018-03-28 10:11:46,031 - DEBUG - ['Prof. Plum\n', 'Miss Scarlet\n', 'Col. Mustard\n', 'Al Sweigart\n', 'Robocop']
	 2018-03-28 10:11:46,032 - DEBUG - It would be a pleasure to have the company of
	 2018-03-28 10:11:46,034 - DEBUG - Prof. Plum

	 2018-03-28 10:11:46,036 - DEBUG - at 11010 Memory Lane on the Evening of
	/usr/local/lib/python3.6/site-packages/docx/styles/styles.py:54: UserWarning: style lookup by style_id is deprecated. Use style name as key instead.
	  warn(msg, UserWarning)
	Traceback (most recent call last):
	  File "custom_invite_word_032718_1.py", line 83, in <module>
	    construct_invite_1( doc,guest_list_maker(guest_txt) )
	  File "custom_invite_word_032718_1.py", line 60, in construct_invite_1
	    doc.paragraphs[2].runs[0].style = 'Heading4' # 'at'
	  File "/usr/local/lib/python3.6/site-packages/docx/text/run.py", line 137, in style
	    style_or_name, WD_STYLE_TYPE.CHARACTER
	  File "/usr/local/lib/python3.6/site-packages/docx/parts/document.py", line 76, in get_style_id
	    return self.styles.get_style_id(style_or_name, style_type)
	  File "/usr/local/lib/python3.6/site-packages/docx/styles/styles.py", line 113, in get_style_id
	    return self._get_style_id_from_name(style_or_name, style_type)
	  File "/usr/local/lib/python3.6/site-packages/docx/styles/styles.py", line 143, in _get_style_id_from_name
	    return self._get_style_id_from_style(self[style_name], style_type)
	  File "/usr/local/lib/python3.6/site-packages/docx/styles/styles.py", line 153, in _get_style_id_from_style
	    (style.type, style_type)
	ValueError: assigned style is type PARAGRAPH (1), need type CHARACTER (2)
	MacBook-Air:custom_invite_word sunnyair$

## Wednesday, March 28, 2018 10:25 AM

Printing it out works fine...

It seems that having multiple invites in the same doc really complicates things...

As it stands right now only the first invite gets any styling...  because the positions in the loop are styling the same ones on the first page...

This means we need a loop that counts numbers for each line of the invite within the guest_list loop...

Reference:

https://python-docx.readthedocs.io/en/latest/user/text.html

Specifically:

"Apply character formatting"

I think I have to create "runs" rather than just "paragraphs" where I can apply the style directly to the run, otherwise just creating a paragraph without creating runs leads to the entire paragraph being the entire run.

	MacBook-Air:custom_invite_word sunnyair$ python3 custom_invite_word_032718_1.py "../tests/guests.txt"
	 2018-03-28 10:24:28,278 - DEBUG - Guest list text file input is:  ../tests/guests.txt
	 2018-03-28 10:24:28,305 - DEBUG - The guest list extracted from text is:
	 2018-03-28 10:24:28,305 - DEBUG - ['Prof. Plum\n', 'Miss Scarlet\n', 'Col. Mustard\n', 'Al Sweigart\n', 'Robocop']
	 2018-03-28 10:24:28,306 - DEBUG - It would be a pleasure to have the company of
	 2018-03-28 10:24:28,308 - DEBUG - Prof. Plum

	 2018-03-28 10:24:28,309 - DEBUG - at 11010 Memory Lane on the Evening of
	 2018-03-28 10:24:28,310 - DEBUG - at 11010 Memory Lane on the Evening of
	 2018-03-28 10:24:28,310 - DEBUG - April 1st
	 2018-03-28 10:24:28,310 - DEBUG - April 1st
	 2018-03-28 10:24:28,311 - DEBUG - at 7 o'clock
	 2018-03-28 10:24:28,312 - DEBUG - at 7 o'clock
	 2018-03-28 10:24:28,312 - DEBUG - It would be a pleasure to have the company of
	 2018-03-28 10:24:28,314 - DEBUG - Prof. Plum

	 2018-03-28 10:24:28,316 - DEBUG - at 11010 Memory Lane on the Evening of
	 2018-03-28 10:24:28,316 - DEBUG - at 11010 Memory Lane on the Evening of
	 2018-03-28 10:24:28,317 - DEBUG - April 1st
	 2018-03-28 10:24:28,317 - DEBUG - April 1st
	 2018-03-28 10:24:28,318 - DEBUG - at 7 o'clock
	 2018-03-28 10:24:28,318 - DEBUG - at 7 o'clock
	 2018-03-28 10:24:28,319 - DEBUG - It would be a pleasure to have the company of
	 2018-03-28 10:24:28,321 - DEBUG - Prof. Plum

	 2018-03-28 10:24:28,323 - DEBUG - at 11010 Memory Lane on the Evening of
	 2018-03-28 10:24:28,323 - DEBUG - at 11010 Memory Lane on the Evening of
	 2018-03-28 10:24:28,324 - DEBUG - April 1st
	 2018-03-28 10:24:28,324 - DEBUG - April 1st
	 2018-03-28 10:24:28,326 - DEBUG - at 7 o'clock
	 2018-03-28 10:24:28,326 - DEBUG - at 7 o'clock
	 2018-03-28 10:24:28,326 - DEBUG - It would be a pleasure to have the company of
	 2018-03-28 10:24:28,328 - DEBUG - Prof. Plum

	 2018-03-28 10:24:28,330 - DEBUG - at 11010 Memory Lane on the Evening of
	 2018-03-28 10:24:28,330 - DEBUG - at 11010 Memory Lane on the Evening of
	 2018-03-28 10:24:28,331 - DEBUG - April 1st
	 2018-03-28 10:24:28,331 - DEBUG - April 1st
	 2018-03-28 10:24:28,333 - DEBUG - at 7 o'clock
	 2018-03-28 10:24:28,333 - DEBUG - at 7 o'clock
	 2018-03-28 10:24:28,333 - DEBUG - It would be a pleasure to have the company of
	 2018-03-28 10:24:28,335 - DEBUG - Prof. Plum

	 2018-03-28 10:24:28,336 - DEBUG - at 11010 Memory Lane on the Evening of
	 2018-03-28 10:24:28,336 - DEBUG - at 11010 Memory Lane on the Evening of
	 2018-03-28 10:24:28,337 - DEBUG - April 1st
	 2018-03-28 10:24:28,337 - DEBUG - April 1st
	 2018-03-28 10:24:28,339 - DEBUG - at 7 o'clock
	 2018-03-28 10:24:28,339 - DEBUG - at 7 o'clock
	MacBook-Air:custom_invite_word sunnyair$

where

	2018-03-28 10:24:28,330 - DEBUG - at 11010 Memory Lane on the Evening of # the paragraph
	2018-03-28 10:24:28,330 - DEBUG - at 11010 Memory Lane on the Evening of # the run[0]

are identical

## Wednesday, March 28, 2018 11:10 AM

We need to use a for loop that increments the positions by a total of 5...

For the very first invite n = 0 * 5 which means no incrementing...

For the 2nd invite n = 1 * 5 which means a five point increment...

## Wednesday, March 28, 2018 11:17 AM

Have to put the position numbering loop as the parent with the name loop as the child.

## Wednesday, March 28, 2018 11:34 AM

	ValueError: assigned style is type PARAGRAPH (1), need type CHARACTER (2)

## Wednesday, March 28, 2018 11:35 AM

[Set paragraph font in python-docx](https://stackoverflow.com/questions/27884703/set-paragraph-font-in-python-docx)

This is how to set the Normal style to font Arial and size 10pt.

	from docx.shared import Pt

	style = document.styles['Normal']
	font = style.font
	font.name = 'Arial'
	font.size = Pt(10)

And this is how to apply it to a paragraph.

	paragraph.style = document.styles['Normal']

Using the current latest version of python-docx (0.8.5)

## Wednesday, March 28, 2018 11:48 AM

Setting base styles

	style.base_style
	logging.debug( style.base_style )
	style.base_style = styles['Normal']
	logging.debug( style.base_style )
	style.base_style.name
	logging.debug( style.base_style.name )

where style is a previously created style object...

	style_citation_1 = styles.add_style('Citation', WD_STYLE_TYPE.PARAGRAPH)
	logging.debug( "The style name created is:  %s" % (style_citation_1.name) )
	logging.debug( "The style type created is:  %s" % (style_citation_1.name) )
	logging.debug( style_citation_1.type )

	style_citation_1.base_style
	logging.debug( style_citation_1.base_style )
	style_citation_1.base_style = styles['Normal']
	logging.debug( style_citation_1.base_style )
	style_citation_1.base_style.name
	logging.debug( style_citation_1.base_style.name )

Another example of creating a new style:

	style_citation_1 = styles.add_style('Citation', WD_STYLE_TYPE.PARAGRAPH)
	logging.debug( "The style name created is:  %s" % (style_citation_1.name) )
	logging.debug( "The style type created is:  %s" % (style_citation_1.name) )
	logging.debug( style_citation_1.type )

## Wednesday, March 28, 2018 11:55 AM

using wrong style type

	2018-03-28 11:54:56,683 - DEBUG - Styles applied
	 2018-03-28 11:54:56,683 - DEBUG - at
	 2018-03-28 11:54:56,683 - DEBUG - 11010 Memory Lane on the Evening of
	Traceback (most recent call last):
	  File "custom_invite_word_032718_1.py", line 116, in <module>
	    construct_invite_1( doc,guest_list_maker(guest_txt) )
	  File "custom_invite_word_032718_1.py", line 86, in construct_invite_1
	    doc.paragraphs[2 + counter].runs[1].style = doc.styles['GuestItalics'] # '11010 Memory Lane on the Evening of'
	  File "/usr/local/lib/python3.6/site-packages/docx/text/run.py", line 137, in style
	    style_or_name, WD_STYLE_TYPE.CHARACTER
	  File "/usr/local/lib/python3.6/site-packages/docx/parts/document.py", line 76, in get_style_id
	    return self.styles.get_style_id(style_or_name, style_type)
	  File "/usr/local/lib/python3.6/site-packages/docx/styles/styles.py", line 111, in get_style_id
	    return self._get_style_id_from_style(style_or_name, style_type)
	  File "/usr/local/lib/python3.6/site-packages/docx/styles/styles.py", line 153, in _get_style_id_from_style
	    (style.type, style_type)
	ValueError: assigned style is type PARAGRAPH (1), need type CHARACTER (2)
	MacBook-Air:custom_invite_word sunnyair$

using:

	style_italic_emph_1 = styles.add_style('GuestItalics', WD_STYLE_TYPE.PARAGRAPH) # use 'GuestItalics' as the name when applying below!

using:

	WD_STYLE_TYPE.PARAGRAPH

seems incorrect

we need character type

[WD_STYLE_TYPE](https://python-docx.readthedocs.io/en/latest/api/enum/WdStyleType.html?highlight=character%20type)

* Specifies one of the four style types: paragraph, character, list, or table.

## Wednesday, March 28, 2018 12:21 PM

This style became optional for the paragraph because I could do it line by line instead...  and still retain the default Word style I used earlier...

	## paragraph styling

	# center_style_1 = doc.styles.add_style('Center1', WD_STYLE_TYPE.PARAGRAPH)  # use 'Center1' as the name when applying below!  also registers it in Word doc styles list
	# center_para_1 = center_style_1.paragraph_format
	# center_para_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
	# # center_para_1.base_style = styles['Subtitle']
	# center_para_1 = center_style_1.font
	# center_para_1.name = 'Calibri'
	# center_para_1.size = Pt(14)
	# center_para_1.italic = True

Never mind, I can still use this style on the paragraphs with runs that have their own styling...

The run styling overrides the paragraph styling...

## Wednesday, March 28, 2018 12:29 PM

The space after the guest name could be fixed by removing the `\n` character that occurs during `readlines()` method usage...



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


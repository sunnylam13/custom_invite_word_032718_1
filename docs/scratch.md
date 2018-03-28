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


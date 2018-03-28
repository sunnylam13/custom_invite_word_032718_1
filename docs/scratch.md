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


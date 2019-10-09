# test-file-generator
 A tool to create files based on web page

## Usage:
  TestFileGenerator.exe --InputFile filenames.txt --Num 200 --OutputDir C:\temp\docs --URL http://en.wikipedia.org/wiki/special:Random

  -i, --InputFile	-	Required. Input file containing list of filenames. Do NOT include extensions.

  -o, --OutputDir	-	Required. Directory to output the documents.

  -n, --Num	-	The number of files to generate. (Default: 200)

  -u, --URL	-	URL of Website to copy (Default: http://en.wikipedia.org/wiki/special:Random)

  --help	-	Display this help screen.

  --version	-	Display version information.

## Background:                                                                                      
While designing a CTF type infrastructure for a Cybersecurity competition, I needed              
to create several network shares to emulate what one might find in a corporate environment.      
I wanted to add some actual files to the shares, but didn't want to manually create them, or     
use real documents.                                                                              
                                                                                                 
This project expands David Palfery's "Test Document Generator" console application               
to accept a file containing potential filenames, a number of documents to generate,              
a URL to copy, and an output directory.       

## Acknowledgments:
David Palfery - https://blogs.perficient.com/2013/04/16/test-document-generator/

## Todo:
Implement functionality to generate more than just Word documents
Automagically create directory structure and save documents throughout
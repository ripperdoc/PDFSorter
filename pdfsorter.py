#!/usr/bin/env python
# coding=utf-8

import sys, os
import argparse, shlex, subprocess, re, unicodedata
from time import strptime, localtime, strftime, mktime
from pyPdf import PdfFileReader, PdfFileWriter
from itertools import izip_longest

def out(s): #############################################################################
  """Print to std out, and if background flag set, to a Growl notification"""
  print s
  if args.background: 
    os.system("""/usr/local/bin/growlnotify -s -i pdf<<END
    %s""" % s)

debug_buffer = ''
def debug(s):
  global debug_buffer
  debug_buffer += s
  debug_buffer += '\n'

def uni_raw(s): #########################################################################
  """Helper function to print unicode chars in a string without encoding them"""
  print s, len(s), type(s)
  l = []
  for c in s:
    l.append(hex(ord(c)))
  print l, len(l)

# Important directories
pdf_tempdir = "/temp/searchable"
pdf_sorted_dir = "/Users/martin/Documents/Scans"
pdf_uncategorized = "/Users/martin/Documents/Scans/Uncategorized"

# The string that identifies PDF text in OS X mdimport command
regex_contents = re.compile(r"kMDItemTextContent = (.*);")

# For each language to support, create the 12 month list and add it to below statements
months_en = ['January','February','March','April','May','June','July','August','September','October','November','December']
months_se = ['januari','februari','mars','april','maj','juni','juli','augusti','september','oktober','november','december']
month_prefixes = [m[0:3].lower() for m in months_en] + [m[0:3].lower() for m in months_se]

# Create a minimal and unique set of matching characters to identify months
months_set = set()
all_months = months_en + months_se
for m in all_months:
  m = m.lower()
  pat = "%s" % m[0:3]
  if len(m)>3:
    pat += "(?:%s)?" % m[3:]
  months_set.add(pat)
months_reg = r"|".join(months_set)

# For more accurate year regex recognition, we will only consider years between
# 1970 and current year, as valid dates for scanned documents
# Below pattern will automatically match up to current year, to make this future proof
cur_year = strftime('%y',localtime())
pattern_year = r'(?:1 ?9)? ?[789] ?[o\d]|(?:2 ?[o0])? ?[o0-%i] ?[o\d]|(?:2 ?[o0])? ?%s ?[o0-%s]' % (int(cur_year[0])-1, cur_year[0], cur_year[1])
# Three formats for dates supported, more can be added
# %y %m %d
# .re match groups(1,2,3,4)
#1999-10-21, 1999 10 21, 1999.10.21, 19991021
#99-10-21, 99 10 21, 99.10.21, 991021
#date_ymd = r'\D(?:19|20)?([0189]\d)(?P<sep1>[- ./])*([01]\d)(?P=sep1)([0-3]\d)\D'
# Second alt, does not assume separators are identical between y-m and m-d
date_ymd = r'[\D]('+pattern_year+')([-. ]?)([o01] ?\d)[ -.]?([o0-3] ?\d)[\D]'

# %d %m %y
# .re match groups(5,6,7,8)
#21/10/1999, 21.10.1999, 21.10.99, 21-10-99
#date_dmy = r'\D([0-3]\d)(?P<sep2>[- ./])+([01]\d)(?P=sep2)(?:19|20)?([0189]\d)\D'
# Second alt, does not assume separators are identical between y-m and m-d
date_dmy = r'[\D]([0-3] ?\d)([- ./]?)([01] ?\d)[- ./]?('+pattern_year+')[\D]'

# %d %b %y
# .re match groups(9,10,11,12)
#21 okt 1999, 21 oktober 1999, 21-oct-1999
#21OCT99, 21 OCT 99
# Also 1 okt 99, e.g. one date number only
#date_dby = r'\D([0-3]?\d)(?P<sep3>[- ./])*('+months_reg+')(?P=sep3)(?:19|20)?([0189]\d)\D'
# Second alt, does not assume separators are identical between y-m and m-d
date_dby = r'[\D]([0-3]? ?\d)([- ./]?)('+months_reg+')[- ./]?('+pattern_year+')[\D]'

# November 27, 2010
date_bdy = r'[\D]('+months_reg+')( ?)([0-3]? ?\d)[, ]+('+pattern_year+')[\D]'

#Not supported yet
#Jun 24 09:30:41 BST 2008

# Some dates are likely to show up regularly in documents but not be the letter date,
# such as the birth date. Add it to this list to not accept these as valid dates
invalid_dates = [strptime('83-02-25','%y-%m-%d')]

# Merge all regexps into one big
regex_literal_date = re.compile(date_dby + "|" + date_bdy, re.I)
regex_date = re.compile(date_ymd + "|" + date_dmy, re.I)

regex_our_date_prefix = re.compile(r'^((19|20)\d\d-[01]\d-[0-3]\d|no_date)_')

# Read the sort dir on folders to find keywords
# Each folder name not directly below the sort dir will count as keyword
# sort_dir
# .. dir1
# .. .. keyword1
# .. dir2
# .. .. keyword2
# etc
keywords = {}
for path, dirs, files in os.walk(unicode(pdf_sorted_dir)):
  ppath, parent = os.path.split(path)
  if ppath == pdf_sorted_dir: # parent is in the sort root, so list of dirs is keywords
    # In OS X filenames, unicode is decomposed, meaning letter and diacritic is 
    # separated. Below line will put them back together, e.g. a + ¨ = ä
    keywords[parent] = [unicodedata.normalize('NFC',d) for d in dirs]

default_prio = 3
prios = {u'American Express':5, 'Skatteverket': 5, 'Nordea':2}
    
flat_p = ''
patterns = []
for k,vals in keywords.iteritems():
  for val in vals:

    if val in prios:
      prio = prios[val]
    else:
      prio = default_prio
    newk = os.path.join(pdf_sorted_dir, k, val)
    val = val.replace(u' ','') #Remove spaces, as we will add arbitrary number of spaces below
    # Make into unicode so that we can correctly put a space between each
    # character - if not, we'd be putting space between each _byte_
    p = ur'\s*'.join(val)
    flat_p += p+'|'
    patterns.append((re.compile(p, re.I | re.U),prio,newk))
patterns = sorted(patterns, key=lambda pattern: pattern[1])
print flat_p
exit()

# Parse input arguments
parser = argparse.ArgumentParser(description="PDF sorters", version=0.1)
parser.add_argument('input', nargs='+', help='files or folders to sort')
parser.add_argument('-n', '--noocr', action='store_true', help='do not run OCR')
parser.add_argument('-d', '--debug', action='store_true', help='show debug output')  
parser.add_argument('-s', '--split', type=int, default=0, help='for each input file, split at every n pages, then sort')
parser.add_argument('-b', '--background', action='store_true', help='background running, output with Growl')
parser.add_argument('-i', '--interactive', action='store_true', help='will prompt user before changing a file')
parser.add_argument('-r', '--recursive', action='store_true', help='will traverse a directory if it\'s the first input')

args = parser.parse_args()

def main(argv): 
  """Run script"""

  def which_matches(patterns, s): ########################################################
    """Checks which of the patterns in provided dictionary that matches string s, return when matched"""
    for pattern,_,key in patterns:
      m = pattern.search(s)
      debug("Match %s = %s" % (pattern.pattern,  m))
      if m is not None:
        return key
    return None

  def get_pdf_contents(pdffile): #########################################################
    """Reads out the text contents of provided PDF file"""
    inputpdf = PdfFileReader(file(pdffile,"rb"))
    #if args.debug: print "PDF Has %i pages" % inputpdf.getNumPages()
    contents = inputpdf.getPage(0).extractText()
    if len(contents.strip())==0:
    	return None
    else:
    	return contents;
    # Run mdimport to parse out PDF contents effectively (only on OS X)
    """sub = subprocess.Popen(shlex.split('/usr/bin/mdimport -d2n "%s"' % pdffile), \
      stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    contents = sub.communicate()[1]
    m = regex_contents.search(contents)
    if m:
      contents = m.group(1)
      # mdimport has different unicode encoding, convert back so Python understands
      contents = contents.replace(r'\U',r'\u').decode('raw_unicode_escape')
    if m is None or len(contents.strip())==0:
      return None
    else:
      return contents"""
    
  def parse_pdf(file, fname, mod_time, contents): ########################################  
    """Parse the pdf to extract based on keyword and date, and return suggested destination
    path and file name""" 
    debug(contents)
    if contents == None:
      return None
    # Find which keyword matches the content (first match only)
    keyword = which_matches(patterns, contents)

    date = "no_date"
    #In debug, show the date patterns used
    debug(regex_literal_date.pattern + '\n') #literal_date=dates with months as words
    debug(regex_date.pattern + '\n')

    # Iterate sequentially (and lazily, e.g. don't match whole string at first) through
    # both literal and normal dates. E.g. compare first literal date and first normal date
    # found side by side before proceed to next tuple of matches.
    for literal_date_match, date_match in izip_longest( \
      regex_literal_date.finditer(contents), regex_date.finditer(contents)):
      
      date,y,m,d = 'invalid_date','','','' # Reset

      if literal_date_match != None:
        if literal_date_match.group(1) != None: #d b y
          d = literal_date_match.group(1).replace(u' ','').lower().replace(u'o','0')
          # All langs first 3 letters in months, lookup position and modulo 12 for
          # month ordinal
          m = (month_prefixes.index(literal_date_match.group(3)[0:3].lower()) % 12) + 1
          y = literal_date_match.group(4).replace(u' ','').lower().replace(u'o','0')[-2:]
          # Although we parse as %d %b %y originally, we have replaced it to %d %m %y
          parsed_format = '%d %m %y'  # We have changed %b to %m
        elif literal_date_match != None and literal_date_match.group(5) != None: # b d y
          m = (month_prefixes.index(literal_date_match.group(5)[0:3].lower()) % 12) + 1
          d = literal_date_match.group(7).replace(u' ','').lower().replace(u'o','0')
          # All langs first 3 letters in months, lookup position and modulo 12 for
          # month ordinal
          y = literal_date_match.group(8).replace(u' ','').lower().replace(u'o','0')[-2:]
          parsed_format = '%m %d %y' # We have changed %b to %m
        try:
          tstruct = strptime('%s %s %s' % (y,m,d),'%y %m %d')
          if tstruct not in invalid_dates:
            date = strftime('%Y-%m-%d',tstruct)
        except ValueError:
          pass
        debug('Found y%s m%s d%s from %s using %s, date is: %s' % \
            (y,m,d,literal_date_match.group(0),parsed_format,date))  
      
      if date_match != None and date=='invalid_date':
        if date_match != None and date_match.group(1) != None: # y m d
          y = date_match.group(1).replace(u' ','').lower().replace(u'o','0')[-2:]
          m = date_match.group(3).replace(u' ','').lower().replace(u'o','0')
          d = date_match.group(4).replace(u' ','').lower().replace(u'o','0')
          parsed_format = '%y %m %d'
        elif date_match != None and date_match.group(5) != None: # d m y
          d = date_match.group(5).replace(u' ','').lower().replace(u'o','0')
          m = date_match.group(7).replace(u' ','').lower().replace(u'o','0')
          y = date_match.group(8).replace(u' ','').lower().replace(u'o','0')[-2:]
          parsed_format = '%d %m %y'

        try:
          tstruct = strptime('%s %s %s' % (y,m,d),'%y %m %d')
          if tstruct not in invalid_dates:
            date = strftime('%Y-%m-%d',tstruct)
        except ValueError:
          pass
        debug('Found y%s m%s d%s from %s using %s, date is: %s' % \
            (y,m,d,date_match.group(0),parsed_format,date))
      
      if date!='invalid_date':
        break #We found a date, leave loop
  
    #print "Dated %s%s, scanned %s" % \
    #  (date, debug_s, strftime('%Y-%m-%d', mod_time))
    if keyword is None:
      destination = os.path.join(pdf_uncategorized,'%s_%s.pdf' % (date, fname))
      safe_content = contents.strip().encode('UTF-8')
      if len(safe_content) > 300:
        print safe_content[0:300] + " [...contd.]"
      else:
        print safe_content
    else:
      # Encode to file system as the 'keyword' variable is pure unicode
      destination = os.path.join(keyword.encode(sys.getfilesystemencoding()), \
        '%s_%s.pdf' % (date, fname))
    return destination

  def handlePdf(curfile, fname, ext, mod_time): ###############################################
    """Take action on a PDF, e.g. OCR it if requested, parse it and move it to new location"""
    title = '\n---%s%s (scanned %s)' % (fname[:49],ext,strftime('%Y-%m-%d', mod_time))
    print title.ljust(81,"-")
    
    ispdf = ext.lower()=='.pdf'
    #print "not args.noocr=%s, not ispdf=%s, get_pdf_contents(curfile) == None = %s" % \
    # (not args.noocr, not ispdf, get_pdf_contents(curfile) == None)
    if not args.noocr and (not ispdf or get_pdf_contents(curfile) == None):
      ocrd_file = os.path.join(pdf_tempdir, fname+".pdf") # Assume same in different folder
      # If not a pdf, OCRed file name may be random and we need to detect diff in dir
      if not ispdf: 
        files_in_dir_before = set(os.listdir(pdf_tempdir))
      cmd = """osascript<<END
      tell application "Adobe Acrobat Pro"
        activate
        set newpath to (POSIX file "%s")
        open newpath
      end tell
      tell application "System Events"
        tell application process "Acrobat"
          click the menu item "OCR This" of menu 1 of menu item "Action Wizard" of the menu "File" of menu bar 1
          repeat until exists (window "OCR This")
          end repeat
          click button "Close" of window "OCR This" 
          click the menu item "Close" of the menu "File" of menu bar 1
        end tell
      end tell
      return""" % curfile #End with return to silence AppleScript output
      os.system(cmd)
      
      if not ispdf:
        addedfile = [f for f in os.listdir(pdf_tempdir) if f not in files_in_dir_before]
        if len(addedfile) > 1:
          exit("%s have multiple added files, cannot tell which one Acrobat just created: %s" % (pdf_tempdir, addedfile))
        else:
          addedfile = addedfile[0]
          #Rename the temp file to the newfile we'd like it to be
          os.rename(os.path.join(pdf_tempdir, addedfile), ocrd_file)
    elif not ispdf: # Can't do non-PDFs without OCR!
      out('Ignoring: cannot handle non-PDF with No OCR setting ON')
      return
    else:
      ocrd_file = curfile
      # The file may have been processed before, remove the date prefix if
      # any to avoid making longer and longer name
      fname = regex_our_date_prefix.sub('',fname,1)

    contents = get_pdf_contents(ocrd_file)
    destination = parse_pdf(ocrd_file, fname, mod_time, contents)
    
    if destination is None:
      out('Error: no text content in %s, ignoring' % ocrd_file)
    elif ocrd_file==destination:
    	out('%s -> Already there' % ocrd_file) 
    else:
      # Pick a unique destination to avoid overwriting
      j=0
      while os.path.exists(destination):
        j+=1
        destination = "%s.%i" % (destination,j)
      out('%s -> %s' % (ocrd_file, destination))
      if args.interactive:
        answer = raw_input("Proceed? [y/n/d(ebug)]")
      else:
        answer = 'y'
      if answer.startswith('y') and not args.debug:
      	os.rename(ocrd_file,destination)
      elif answer.startswith('d'):
      	print debug_buffer
      	# TODO avoid repetition, create small function
      	print '%s -> %s' % (ocrd_file, destination)
      	answer = raw_input("Proceed? [y/n]")
      	if answer.startswith('y') and not args.debug:
      	  os.rename(ocrd_file,destination)
    return

  # If first input is a dir, enumerate the pdf files in it (and ignore rest of input)
  extensions = ['.pdf','.jpg','.jpeg','.png','.gif','.bmp']
  if os.path.isdir(args.input[0]):
    parsefiles = []
    for path, dirs, files in os.walk(args.input[0]):
      parsefiles.extend([os.path.join(path,f) for f in files if os.path.splitext(f)[1].lower() in extensions])
      if not args.recursive:
        break # Only do at top level if recursive is not on
  # Or use the input list of files directly
  else:
    parsefiles = args.input
    
  for curfile in parsefiles:
    curfile = os.path.abspath(curfile)
    fname = os.path.split(curfile)[1] # pick file name part
    (fname,ext) = os.path.splitext(fname) # remove extension
    # Get modified time on the file
    mod_time = localtime(os.stat(curfile).st_mtime)
    
    # Split PDFs into parts before handling them, if requested
    if args.split > 0:
      inputpdf = PdfFileReader(file(curfile,"rb"))
      mtime = (mktime(mod_time),mktime(mod_time))
      if inputpdf.numPages > args.split: # Only process if we expect to split
        outputpdf = PdfFileWriter()
        curPage = 0
        while curPage<inputpdf.numPages:
          print "Adding page %s" % curPage
          outputpdf.addPage(inputpdf.getPage(curPage))
          if (curPage+1) % args.split == 0: #Time to split, last page before split
            new_fname = '%s_pp%i-%i' % (fname,curPage+1-curPage % args.split,curPage+1)
            curfile = os.path.join(pdf_tempdir,new_fname+ext)
            outputStream = file(curfile,"wb")
            outputpdf.write(outputStream)
            outputStream.close()
            os.utime(curfile, mtime) # Set same mod time as parent file
            handlePdf(curfile, new_fname, ext, mod_time)
            outputpdf = PdfFileWriter()
          curPage += 1
        if curPage % args.split != 0: # NOT last page before split, so uneven page(s) left
          new_fname = '%s_pp%i-%i' % (fname,curPage+1-curPage % args.split,curPage+1)
          curfile = os.path.join(pdf_tempdir,new_fname+ext)
          outputStream = file(curfile,"wb")
          outputpdf.write(outputStream)
          outputStream.close()
          os.utime(curfile, mtime) # Set same mod time as parent file
          handlePdf(curfile, new_fname, ext, mod_time)
    else:
      handlePdf(curfile, fname, ext, mod_time)
  return
  
if __name__ == '__main__': sys.exit(main(sys.argv))
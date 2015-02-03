#!/usr/bin/env python
#       This program is free software; you can redistribute it and/or modify
#       it under the terms of the GNU General Public License as published by
#       the Free Software Foundation; either version 2 of the License, or
#       (at your option) any later version.
#       
#       This program is distributed in the hope that it will be useful,
#       but WITHOUT ANY WARRANTY; without even the implied warranty of
#       MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#       GNU General Public License for more details.
#       
#       You should have received a copy of the GNU General Public License
#       along with this program; if not, write to the Free Software
#       Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
#       MA 02110-1301, USA.
#       

# -*- coding: utf-8 -*-
import urllib
import urllib2
import re
from xlwt import *

SPREADSHEET_NAME = 'Amazon2.xls'

class Star:
	# 48 potential fields.
	#self.attrib = {'jack': 4098, 'sape': 4139}
	def __init__(self, index):
		self.att = {'numberid': index}
		self.headings = ['numberid', 'name', 'userid', 'profileurl', 'realname', 'vine', 'totalreviews', 'topreviewer', 'email', 'halloffame',
		'2000', '2001', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', 'votes', 
		'helpful', 'ratio', 'location', 'tags', 'info']
		for year in range(2000, 2013):
			self.att[str(year)] = 0 # zero all the years to start.
		
	def add(self, name, value):
		# self.att[name] = product.encode_utf8(value)
		self.att[name] = value
		
	def get( self, name ):
		return self.att[name]
		
	def getKeys(self):
		return self.headings
		
	def getProfileURL(self):
		try:
			return self.att['profileurl']
		except KeyError:
			return None
		
	def toStr(self):
		if (DEBUG):
			for key in self.att.keys():
				print key + " = " + str(self.att[key])
		else:
			for key in self.att.keys():
				print key + " = " + self.att[key]
			
	def writeSS(self, sheet, row):
		style = XFStyle()
		fnt = Font()
		fnt.name = 'Arial'
		style.font = fnt
		style_link = XFStyle()
		fnt_link = Font()
		fnt_link.name = 'Arial'
		fnt_link.colour_index = 0x4
		style_link.font = fnt_link
		
		index = 0
		for heading in self.headings:
			try:
				sheet.write(row, index, self.att[heading], style)
			except KeyError:
				print "missing " + heading
			index += 1
		#ws.write(row, 2, changes, style) # this for standard text to a cell
		#ws.write(row, 2, Formula('HYPERLINK("' + changes + '";"' + changes + '")'), style_link)
		return
		
	# This method takes teh URL for the products the reviewer has reviewed and scrapes those pages.
	# each page will will contain all the products the reviewer has reviewed, one for each line
	# and the page is named after the reviewer. Since this is expensive to query each page 
	# we will perform checkpointing so we don't have to redo all the entries if we loose connection.
	def getMyProductReviewPages(self, spreadsheet, fileName):
		# make a page for the reviewer with their name
		try:
			productSheet = spreadsheet.add_sheet(self.att['name'])
			product.getProductReviews(self.att['name'], self.att['userid'], self.att['reviewurl'], productSheet)
			spreadsheet.save(fileName)
		except UnicodeDecodeError:
			print "Error encoding values within\n'" + self.att['name'] + "' in reviewURL:\n" +  self.att['reviewurl']
		except KeyError: # warn the user that either name or reviewurl wasn't found.
			print "Key error: either the reviewer's name or 'reviewurl' wasn't found in the reviewer's main page."
			return # we move on to the next one.

# Formulates the parameters into a valid Widipedia API
# call
# Inputs: parameters - dictionary of all the parameters needed to fufil the call
# in standard UTF-8 format.
# Returns: the page response from Wikipedia.
import os.path
#import urllib2

COOKIEFILE = 'cookies.lwp'
# the path and filename to save your cookies in

cj = None
ClientCookie = None
cookielib = None

# Let's see if cookielib is available
try:
    import cookielib
except ImportError:
    # If importing cookielib fails
    # let's try ClientCookie
    try:
        import ClientCookie
    except ImportError:
        # ClientCookie isn't available either
        urlopen = urllib2.urlopen
        Request = urllib2.Request
    else:
        # imported ClientCookie
        urlopen = ClientCookie.urlopen
        Request = ClientCookie.Request
        cj = ClientCookie.LWPCookieJar()

else:
    # importing cookielib worked
    urlopen = urllib2.urlopen
    Request = urllib2.Request
    cj = cookielib.LWPCookieJar()
    # This is a subclass of FileCookieJar
    # that has useful load and save methods
    
# RefererUrl allows us to state the page that we came from.
def query_URL(url, refererUrl=None):
	# we do this to initialize the return page so if the query fails because 
	# of URL errors, it can continue to the next article.
	the_page = ""
	#if (DEBUG): # add this because network is unreliable with kids watching Netflix.
	#	f = open('return_page.html')
	#	the_page = f.read()
	#	f.close()
	#	return the_page
	#try:
	#response = urllib2.urlopen(url)
	#the_page = response.read()
	#####################################
	#cj = cookielib.CookieJar()
	#opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))
	#r = opener.open(url)
	#the_page = r.read()
	########################################
	#if cj is not None:
	# we successfully imported
	# one of the two cookie handling modules
	if os.path.isfile(COOKIEFILE):
		# if we have a cookie file already saved
		# then load the cookies into the Cookie Jar
		#print "loading cookie file"
		cj.load(COOKIEFILE)
	# Now we need to get our Cookie Jar
	# installed in the opener;
	# for fetching URLs
	#if cookielib is not None:
		# if we use cookielib
		# then we get the HTTPCookieProcessor
		# and install the opener in urllib2
	opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))
	urllib2.install_opener(opener)
	#print "CookieProcessor invoked."
	#else:
	#	# if we use ClientCookie
	#	# then we get the HTTPCookieProcessor
	#	# and install the opener in ClientCookie
	#	opener = ClientCookie.build_opener(ClientCookie.HTTPCookieProcessor(cj))
	#	ClientCookie.install_opener(opener)
	#	#theurl = 'http://www.google.co.uk/search?hl=en&ie=UTF-8&q=voidspace&meta='
	#	# an example url that sets a cookie,
	#	# try different urls here and see the cookie collection you can make !

	txdata = None
	# if we were making a POST type request,
	# we could encode a dictionary of values here,
	# using urllib.urlencode(somedict)

	txheaders =  {'User-agent' : 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'}
	# fake a user agent, some websites (like google) don't like automated exploration

	try:
		#req = Request(theurl, txdata, txheaders)
		req = Request(url, txdata, txheaders)
		if (refererUrl != None):
			req.add_header('Referer', refererUrl)
		# create a request object

		handle = urlopen(req)
		# and open it to return a handle on the url
		the_page  = handle.read()
	except IOError, e:
		print 'We failed to open "%s".' % repr(e)
		if hasattr(e, 'code'):
			print 'We failed with error code - %s.' % e.code
		elif hasattr(e, 'reason'):
			print "The error object has the following 'reason' attribute :"
			print e.reason
			print "This usually means the server doesn't exist,",
			print "is down, or we don't have an internet connection."
		return "none"

	#else:
		#print 'Here are the headers of the page :'
		#print handle.info()
		# handle.read() returns the page
		# handle.geturl() returns the true url of the page fetched
		# (in case urlopen has followed any redirects, which it sometimes does)

		#print

	#print 'These are the cookies we have received so far :'
	#for index, cookie in enumerate(cj):
	#	print index, '  :  ', cookie
	#print "\n=== end of cookie transcript ===\n\n"
	cj.save(COOKIEFILE)                     # save the cookies again
	##########################################################
	#cookie_handler = urllib2.HTTPCookieProcessor( cookies )
	#redirect_handler = HTTPRedirectHandler()
	#opener = urllib2.build_opener(redirect_handler, cookie_handler)
	#response = opener.open(url)
	#the_page  = response.read()
	#except Exception, e1:
	#	print "Error reading from URL ", url, "::",
	#	if hasattr(e1, 'code'):
	#		print 'We failed with error code - %s.' % e1.code
	#	elif hasattr(e1, 'reason'):
	#		print "The error object has the following 'reason' attribute :"
	#		print e1.reason
	return the_page


	
def remove_html_tags(data):
    p = re.compile(r'<.*?>')
    return p.sub('', data)



def write_ss_headings(ws):
	fnt = Font()
	fnt.name = 'Arial'
	fnt.bold = True
	borders = Borders()
	borders.bottom = 1
	style = XFStyle()
	style.font = fnt
	style.borders = borders
	i = 0
	s = Star(0)
	for heading in s.getKeys():
		ws.write(0, i, heading, style)
		i += 1

def getReviewers( page, index ):
	reviewers = []
	myIndex = index
	myLines = page.split( "<tr id=\"reviewer" )[1:]
	for line in myLines:
		# create a reviewer 
		reviewer = Star( myIndex )
		# print line
		# parse out the UserID
		userId = line.split( "http://www.amazon.com/gp/pdp/profile/" )[1].split( "/ref" )[0]
		print "=>" + userId,
		reviewer.add( "userid", userId )
		# oiy vey this is bad...
		userURL = line.split( "http://www.amazon.com/gp/pdp/profile/" )[1].split( "_pic" )[0]
		userURL = "http://www.amazon.com/gp/pdp/profile/" + userURL + "_name"
		print "=>" + userURL,
		reviewer.add( "profileurl", userURL )
		name = line.split( "_name\"><b>" )[1].split( "</b" )[0]
		print "=>" + name,
		reviewer.add( "name", name )
		totalReviews = line.split( "See all " )[1].split( " reviews" )[0]
		print "=>" + totalReviews,
		reviewer.add( "totalreviews", totalReviews )
		percentHelpful = line.split( "crNumPercentHelpful\"> " )[1].split( " </td>" )[0]
		print "=>" + percentHelpful
		reviewer.add( "ratio", percentHelpful )
		reviewers.append( reviewer )
		myIndex += 1
	return reviewers
	
def getReviewersPersonalData( reviewer ):
	page = encode_utf8( query_URL( reviewer.getProfileURL() ) )
	if ( page == "none" ):
		return 0
	print "----------"
	# print page # gets the hold of fame review years.
	years = page.split( '<div class="hallofFameYears">' )
	if (len(years) > 1):
		yearString = years[1].split('</div>')[0]
		years = yearString.split()
		for year in years: # like: Hall of Fame Reviewer - 2000 2001 2003 2004 2005 2006 2007 2008 2009 2010 2011
			if year[0].isdigit():
				reviewer.add( year, '1' )
	
	if (page.find('REAL NAME') > -1): # from the alt attribute.
		reviewer.add('realname', '1')
	else:
		reviewer.add('realname', '0')
	if (page.find('FAME REVIEWER') > -1): # restricted because #1 HaLL OF FAME not HALL OF FAME
		reviewer.add('halloffame', '1')
	else:
		reviewer.add('halloffame', '0')
	if (page.find('VINE VOICE') > -1): 
		reviewer.add('vine', '1')
	else:
		reviewer.add('vine', '0')
	if (page.find('Top Reviewer Ranking:') > -1):
		reviewer.add('topreviewer', '1')
	else:
		reviewer.add('topreviewer', '0')
	if (page.find('E-mail:') > -1):
		reviewer.add('email', '1')
	else:
		reviewer.add('email', '0')
	if (page.find( 'Location:</b>' ) > -1):
		location = page.split( 'Location:</b>' )[1].split( '</div>' )[0]
		reviewer.add('location', location)
	else:
		reviewer.add('location', 'n/a' )
		
	# now to collect the helpful review count
	revsText = page.split( 'Helpful votes received on reviews:' )
	try:
		rvsText  = revsText[1].split( '%</b> (' )[1].split( ')' )[0]
		#print ">****" + rvsText + "****<"
		votes = rvsText.split( ' of ' )
		reviewer.add( "votes", votes[1] )
		reviewer.add( "helpful", votes[0] )
	except IndexError, e:
		reviewer.add( "votes", "n/a" )
		reviewer.add( "helpful", "n/a" )
	#tags
	try:
		ts = page.split( 'Interests' )[1].split( '<div style="margin-top: 2px">' )[1].split( '</div>' )[0]
		#print ">****" + ts + "****<"
		reviewer.add( "tags", ts )
	except IndexError, e:
		reviewer.add( "tags", "n/a" )
	try:
		infos = page.split( 'In My Own Words:' )[1].split( '<div style="margin-top: 2px">' )[1].split( '</div>' )[0]
		info  = ' '.join(infos.split());
		#print ">****" + info + "****<"
		reviewer.add( "info", info )
	except IndexError, e:
		reviewer.add( "info", "n/a" )
	return 1
			
# Why the hell am I doing this??
# Oh, yeah duffus had to put trademarks in the content!
def encode_utf8(text):
	st = ""
	for ch in text:
		if (ord(ch) > 128):
			st += ' '
			continue
		st += ch
	return st
	
DEBUG = False
stateFileName = 'aws_research.save'
limit = 2
count = 0
# The whole thing starts here.
if __name__ == "__main__":
	import doctest
	doctest.testmod()
	# open a spreadsheet
	spreadsheet = Workbook()
	style = XFStyle()
	row = 1
	finished = row
	count = 0;  #Row 1; row 0 is the title.
	if (os.path.isfile( stateFileName )):
		f = open( stateFileName, 'r' )
		leftOffLine = f.readlines()
		row = leftOffLine[0]
		finished = row
		f.close()
		count = row -1
	else:	
		wsheet = spreadsheet.add_sheet("Amazon Top Customer Reviewers")
		write_ss_headings(wsheet)
	
	f = open('urls_to_scrape.lst', 'r')
	lines = f.readlines()
	f.close()
	for line in lines[finished:]:
		print "URL #" + str(row) + " of " + str( len (lines[finished:]))
		if (len(line) > 0):
			#print "===" + line + "==="
			page = encode_utf8( query_URL( line ) )
			if ( page == "none" ):
				# preserve location since we lost connection to the internet. resume on next iteration.
				spreadsheet.save( SPREADSHEET_NAME )
				fState = open( stateFileName, 'w' )
				fState.write( str(row) )
				fState.close()
			# get the reviewers listed on the page as a list of URLs to those pages.
			tmpReviewers = getReviewers( page, row )
			for reviewer in tmpReviewers:
				if (getReviewersPersonalData( reviewer ) == 0):
					spreadsheet.save( SPREADSHEET_NAME )
					fState = open( stateFileName, 'w' )
					fState.write( str(row) )
					fState.close()
				reviewer.writeSS( wsheet, row )
				print "-->printing: reviewer " + reviewer.get( "name" ) + " on row " + str( row )
				row += 1
				spreadsheet.save( SPREADSHEET_NAME )
				if DEBUG and count >= limit:
					break
				count += 1

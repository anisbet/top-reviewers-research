#!/usr/bin/env python
# -*- coding: utf-8 -*-
import urllib
import urllib2
import re
from xlwt import *
#from xlrd import open_workbook, cellname, xldate_as_tuple
#from datetime import datetime, date, time
#import json
import product

SPREADSHEET_NAME = 'Amazon.xls'

class Star:
	# 48 potential fields.
	#self.attrib = {'jack': 4098, 'sape': 4139}
	def __init__(self, index):
		self.att = {'numberid': index}
		self.headings = ['numberid', 'name', 'userid', 'profileurl', 'realname', 'vine', 'topreviewer', 'email', 'halloffame',
		'2000', '2001', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', 'votes', 
		'helpful', 'ratio', 'location', 'tags', 'info']
		for year in range(2000, 2013):
			self.att[str(year)] = 0 # zero all the years to start.
		
	def add(self, name, value):
		self.att[name] = value
		
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
	# we will perform checkpointing so we don't have to redo all the entries if we lose connection.
	def getMyProductReviewPages(self, spreadsheet, fileName):
		# make a page for the reviewer with their name
		try:
			productSheet = spreadsheet.add_sheet(self.att['name'])
			product.getProductReviews(self.att['name'], self.att['reviewurl'], productSheet)
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
def query_URL(url):
	# we do this to initialize the return page so if the query fails because 
	# of URL errors, it can continue to the next article.
	the_page = ""
	#if (DEBUG): # add this because network is unreliable with kids watching Netflix.
	#	f = open('return_page.html')
	#	the_page = f.read()
	#	f.close()
	#	return the_page
	try:
		response = urllib2.urlopen(url)
		the_page = response.read()
	except Exception:
		print "Error reading from URL ", url
	return the_page
	
# Gets the details about the reviewer from the argument HTML <td>. These details are from the main
# star reviewer page. See setReviewDetails() for the data about reviews.
# param: tds - table data tags for each reviewer.
# param: reviewer - Star object for deets to be filled in.
# return:
def setReviewersDetails(tds, reviewer):
	rId = ""
	index = 0
	for td in tds:
		#print "=>" + td + "<=\n\n\n"
		tmpArray = td.split('/profile/')
		if (index == 2):
			if (td.find('REAL NAME') > -1): # from the alt attribute.
				reviewer.add('realname', 1)
			else:
				reviewer.add('realname', 0)
			if (td.find('FAME REVIEWER') > -1): # restricted because #1 HaLL OF FAME not HALL OF FAME
				reviewer.add('halloffame', 1)
			else:
				reviewer.add('halloffame', 0)
			if (td.find('VINE VOICE') > -1):
				reviewer.add('vine', 1)
			else:
				reviewer.add('vine', 0)
			if (td.find('TOP 50 REVIEWER') > -1):
				reviewer.add('topreviewer', 1)
			else:
				reviewer.add('topreviewer', 0)
			if (td.find('EMAIL') > -1):
				reviewer.add('email', 1)
			else:
				reviewer.add('email', 0)
		elif (index == 1):
			# if this is defined then we should strip the tags and sift out a name.
			data = td.split('<div')
			name = remove_html_tags(data[0])
			# get rid of the end of the div tag angle bracket
			name = name[1:].lstrip()
			reviewer.add('name', name)
			# find the links for all their reviews in the nearby link.
			reviewsLink = get_review_link(td)
			# this is not an official heading so it will not be put into the spreadsheet
			# but it can be referenced by the reviewer when we want to get the product reviews.
			reviewer.add('reviewurl', reviewsLink)
		elif (index == 0):
			# get the reviewer's id
			rId = tmpArray[1].split('/')[0]
			rId = rId.split('">')[0]
			reviewer.add('userid', rId) # should contain the users profile id.
			# the profile URL is also in this <td>
			profileURL = 'http://www.amazon.com/gp/pdp/profile/'+rId
			reviewer.add('profileurl', profileURL)
		index += 1
	print "scraped reviewer page '" + reviewer.att['name'] + "'"
	return
	
# This function fires to collect the information from the profile pages
def setReviewersProfile(reviewer):
	url = reviewer.getProfileURL()
	if (url == None):
		print 'reviewers profile URL could not be found.';
		return
	# now get the years they have been reviewing; potentially 2000 - 2012 for Kristen's study anyway
	data = query_URL(url)
	years = data.split('<div class="hallofFameYears">')
	if (len(years) > 1):
		yearString = years[1].split('</div>')[0]
		years = yearString.split()
		for year in years: # like: Hall of Fame Reviewer - 2000 2001 2003 2004 2005 2006 2007 2008 2009 2010 2011
			if year[0].isdigit():
				reviewer.add(year, 1)
	# now get the helpful votes:
	votesRaw = data.split('<span class="label">Helpful votes received on reviews:</span>')
	if (len(votesRaw) > 1):
		voteString = votesRaw[1].split('</span>')[0]
		voteString = remove_html_tags(voteString)
		voteString = voteString.replace("(", " ").replace(")", " ").replace("of", " ")
		#print voteString + "++++++++++++++++",
		votes = voteString.split()
		#print len(votes)
		reviewer.add('votes', votes[2])
		reviewer.add('helpful', votes[1])
		reviewer.add('ratio', votes[0])
	locationRaw = data.split('<b>Location:</b>')
	if (len(locationRaw) > 1):
		locationString = locationRaw[1].split('</div>')[0]
		reviewer.add('location', locationString)
	tagsRaw = data.split('Frequently Used Tags')
	if (len(tagsRaw) > 1):
		tagsString = tagsRaw[1].split('</div>')[0]
		tagsString = remove_html_tags(tagsString)
		tags = tagsString.splitlines()
		authorTags = []
		for tag in tags:
			tag = tag.lstrip()
			tag = tag.rstrip()
			if (tag != ""):
				authorTags.append(tag)
		tagsString = ', '.join(authorTags)
		#print tagsString + "++++++++++++++++"
		reviewer.add('tags', tagsString)
	infoRaw = data.split('In My Own Words:')
	if (len(infoRaw) > 1):
		infoString = infoRaw[1].split('<a href=')[0]
		infoString = remove_html_tags(infoString).splitlines()
		retString = ""
		for line in infoString:
			line = line.lstrip()
			if line == 'Interests' or line.startswith('Frequent'):
				break
			retString = retString + line
		#print retString + "++++++++++++++++"
		reviewer.add('info', retString)

# returns the member's reviews page from the main Star Reviewer's page.
def get_review_link(data):
	hrefs = data.split('<a href="')
	for href in hrefs:
		#print href + "(==\n\n"
		if (href.find('member-reviews') > -1):
			return href[0: href.find('">')]
	return "no review URL found"
	
def remove_html_tags(data):
    p = re.compile(r'<.*?>')
    return p.sub('', data)
	
# Returns an array of star reviewer objects.
# Param: page - html page from Amazon.
# Return: array of reviewers ([Star_reviewer1, Star_reviewer2, ...]).
def getStarReviewers(page):
	reviewers_HTML = page.split('id="halloffameReviewer"')
	reviewers_HTML.pop(0) # get rid of first element it doesn't contain any useful info.
	print str(len(reviewers_HTML)) + " reviewers listed on this page."
	# now we have the page split roughly into reviewers stats let's get deets for each.
	reviewers = []
	reviewNumber = 1
	i = 0 # for debugging exit early.
	for reviewer_HTML in reviewers_HTML:
		# Amazon stores their reviewers in set of tables. Split up the table data.
		myReviewersTags = reviewer_HTML.split('<td') # split on td for each record.
		myReviewersTags.pop(0) # remove the first which doesn't contain any useful data.
		reviewer = Star(reviewNumber)
		reviewNumber += 1
		setReviewersDetails(myReviewersTags, reviewer)
		setReviewersProfile(reviewer)
		reviewers.append(reviewer)
		if (DEBUG and i == 4):
			break
		i += 1
	return reviewers # TODO return object array.
	
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

DEBUG = True

# The whole thing starts here.
if __name__ == "__main__":
	import doctest
	doctest.testmod()
	page = query_URL('http://www.amazon.com/review/hall-of-fame')
	star_reviewers = getStarReviewers(page)
	# open a spreadsheet
	spreadsheet = Workbook()
	style = XFStyle()
	wsheet = spreadsheet.add_sheet("Amazon Star Reviewers")
	write_ss_headings(wsheet)
	index = 1;  #Row 1; row 0 is the title.
	for star in star_reviewers:
		if (DEBUG):
			print star.toStr()
		star.writeSS(wsheet, index)
		index += 1
	spreadsheet.save(SPREADSHEET_NAME)
	for star in star_reviewers:
		star.getMyProductReviewPages(spreadsheet, SPREADSHEET_NAME)
		spreadsheet.save(SPREADSHEET_NAME)

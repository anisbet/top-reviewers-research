#!/usr/bin/env python
# -*- coding: utf-8 -*-
import urllib
import urllib2
import re
from xlwt import *
#from xlrd import open_workbook, cellname, xldate_as_tuple
#from datetime import datetime, date, time
#import json


class Star:
	# 48 potential fields.
	#self.attrib = {'jack': 4098, 'sape': 4139}
	def __init__(self, index):
		self.att = {'numberid': index}
		self.headings = ['numberid', 'name', 'userid', 'profileurl', 'realname', 'vine', 'topreviewer', 'email', 'halloffame',
		'2000', '2001', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', 'votes', 
		'helpful', 'ratio', 'location', 'tags', 'info', 'firstrev', 'totalrev', 'reviewstar', 'reviewtitle'
		'reviewdate', 'reviewyear', 'dayssincelast', 'dayssincefirst', 'votes', 'helpful', 'content', 'characters',
		'product', 'producturl', 'category1', 'category2', 'category3', 'productfirst', 'daysfirst', 'avreview']
		
	def add(self, name, value):
		self.att[name] = value
		
	def getKeys(self):
		return self.headings
		
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

# Formulates the parameters into a valid Widipedia API
# call
# Inputs: parameters - dictionary of all the parameters needed to fufil the call
# in standard UTF-8 format.
# Returns: the page response from Wikipedia.
def query_URL(url):
	# we do this to initialize the return page so if the query fails because 
	# of URL errors, it can continue to the next article.
	the_page = ""
	if (DEBUG): # add this because network is unreliable with kids watching Netflix.
		f = open('return_page.html')
		the_page = f.read()
		f.close()
		return the_page
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
		print "=>" + td + "<=\n\n\n"
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
			reviewer.add('reviewURL', reviewsLink)
		elif (index == 0):
			# get the reviewer's id
			rId = tmpArray[1].split('/')[0]
			rId = rId.split('">')[0]
			reviewer.add('userid', rId) # should contain the users profile id.
			# the profile URL is also in this <td>
			profileURL = 'http://www.amazon.com/gp/pdp/profile/'+rId
			reviewer.add('profileurl', profileURL)
		index += 1
	return

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
	index = 0
	for reviewer_HTML in reviewers_HTML:
		# Amazon stores their reviewers in set of tables. Split upt the table data.
		myReviewersTags = reviewer_HTML.split('<td') # split on td for each record.
		myReviewersTags.pop(0) # remove the first which doesn't contain any useful data.
		reviewer = Star(index)
		index += 1
		setReviewersDetails(myReviewersTags, reviewer)
		reviewers.append(reviewer)
		if (DEBUG):
			break
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

DEBUG = False

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
	index = 1; # top reviewer get 0, next 1 etc.
	for star in star_reviewers:
		if (DEBUG):
			print star.toStr()
		else:
			star.writeSS(wsheet, index)
		index += 1
	spreadsheet.save('Amazon.xls')

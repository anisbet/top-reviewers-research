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
	def __init__(self):
		self.attributes = {}
		
	def add(self, name, value):
		self.attributes[name] = value
		
	def toStr(self):
		for key in self.attributes.keys():
			print key + " = " + self.attributes[key]
			
	def writeSS(self, sheet):
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
	
# Gets the details about the reviewer from the argument HTML <td>
# param: tds - table data tags for each reviewer.
# param: reviewer - Star object for deets to be filled in.
# return:
def setReviewersDetails(tds, reviewer):
	rId = ""
	index = 0
	for td in tds:
		#print "=>" + td + "<=\n\n\n"
		tmpArray = td.split('/profile/')
		if (index == 1):
			# if this is defined then we should strip the tags and sift out a name.
			data = td.split('<div')
			name = remove_html_tags(data[0])
			reviewer.add('name', name)
			# find the links for all their reviews in the nearby link.
			reviewsLink = get_review_link(td)
			reviewer.add('reviewURL', reviewsLink)
		if (index == 0):
			rId = tmpArray[1].split('/')[0]
			rId = rId.split('">')[0]
			reviewer.add('id', rId) # should contain the users profile id.
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
	for reviewer_HTML in reviewers_HTML:
		# Amazon stores their reviewers in set of tables. Split upt the table data.
		myReviewersTags = reviewer_HTML.split('<td') # split on td for each record.
		myReviewersTags.pop(0) # remove the first which doesn't contain any useful data.
		reviewer = Star()
		setReviewersDetails(myReviewersTags, reviewer)
		reviewers.append(reviewer)
		if (DEBUG):
			break
	return reviewers # TODO return object array.
	
def write_ss_headings(ws):
	headings = ['numberid', 'name', 'userid', 'profileurl', 'realname', 'vine', 'topreviewer', 'email', 'halloffame',
		'2000', '2001', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', 'votes', 
		'helpful', 'ratio', 'location', 'tags', 'info', 'firstrev', 'totalrev', 'reviewstar', 'reviewtitle'
		'reviewdate', 'reviewyear', 'dayssincelast', 'dayssincefirst', 'votes', 'helpful', 'content', 'characters',
		'product', 'producturl', 'category1', 'category2', 'category3', 'productfirst', 'daysfirst', 'avreview']
	fnt = Font()
	fnt.name = 'Arial'
	fnt.bold = True
	borders = Borders()
	borders.bottom = 1
	style = XFStyle()
	style.font = fnt
	style.borders = borders
	i = 0
	for heading in headings:
		ws.write(0, i, heading, style)
		i += 1


DEBUG = False

# The whole thing starts here.
if __name__ == "__main__":
	import doctest
	doctest.testmod()
	index = 0; # top reviewer get 0, next 1 etc.
	page = query_URL('http://www.amazon.com/review/hall-of-fame')
	star_reviewers = getStarReviewers(page)
	# open a spreadsheet
	spreadsheet = Workbook()
	style = XFStyle()
	wsheet = spreadsheet.add_sheet("Amazon Star Reviewers")
	write_ss_headings(wsheet)
	for star in star_reviewers:
		if (DEBUG):
			print star.toStr()
		else:
			spreadsheet.save('wiki_data.xls')
	spreadsheet.save('Amazon.xls')

#!/usr/bin/env python
# -*- coding: utf-8 -*-
# gets information from individual product pages. Each product should envoke a save on the spread sheet.

import reviewer
import re
from xlwt import *

class Product:
	
	def __init__(self):
		self.att = {}
		self.productHeadings = [
			# scrape,        scrape,        scrape,       extract 'reviewdate'
			'reviewstar', 'reviewtitle', 'reviewdate', 'reviewyear', # Reviewer's product review page
			'content', 'characters', 'product', 'producturl', # Reviewer's product pages to avoid <read more> link
			'firstrev', 'totalrev', 'votes', 'helpful', #community review of product
			# some pages,     compute,     scrape
			'productfirst', 'daysfirst', 'avreview', # community review.
			#  compute (today - reviewdate), compute (today - firstrev)
			'dayssincelast', 'dayssincefirst', #computed from date of today - reviewdate, today - firstrev
			'category1', 'category2', 'category3'] #community review of product
		self.att['votes']    = 0
		self.att['helpful']  = 0
		self.att['avreview'] = 0
		self.att['category1'] = "N/A"
		self.att['category2'] = "N/A"
		self.att['category3'] = "N/A"
	
	def getHeadings(self):
		return self.productHeadings
		
	def add(self, name, value):
		self.att[name] = value
		
	def get(self, which):
		try:
			return self.att[which];
		except KeyError:
			return None
		
	def toStr(self):
		if (reviewer.DEBUG):
			for key in self.att.keys():
				print key + " = " + str(self.att[key])
		else:
			for key in self.att.keys():
				print key + " = " + str(self.att[key])

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
		for heading in self.productHeadings:
			try:
				sheet.write(row, index, self.att[heading], style)
			except KeyError:
				print "missing " + heading
			index += 1
		#ws.write(row, 2, changes, style) # this for standard text to a cell
		#ws.write(row, 2, Formula('HYPERLINK("' + changes + '";"' + changes + '")'), style_link)
		return

# Writes the headings for each of the reviewer's sheets.
def writeProductHeadings(worksheet):
	fnt = Font()
	fnt.name = 'Arial'
	fnt.bold = True
	borders = Borders()
	borders.bottom = 1
	style = XFStyle()
	style.font = fnt
	style.borders = borders
	i = 0
	product = Product()
	for heading in product.getHeadings():
		worksheet.write(0, i, heading, style)
		i += 1

# called from Star object, it will take the reviewers product url and crawl the links 
# there for all the product s/he has reviewed, making a row entry for each and a 
# saving as we go.
# param reviewer 
# param sheet for reviews.
def getProductReviews(reviewer_name, reviewer_id, reviewURL, ssheet):
	writeProductHeadings(ssheet)
	# here we will get the information we can from the Reviewer's product page:
	allProducts = []
	parsePage(reviewURL, ssheet, allProducts)
	# to get here we have scraped all the data from the reviewer's product page
	# now get the product data from the Products pages.
	href = 'http://www.amazon.com/gp/cdp/member-reviews/' + reviewer_id + "?ie=UTF8&amp;display=public&amp;sort_by=MostRecentReview&amp;page="
	#href = '/gp/cdp/member-reviews/' + reviewer_id + "?ie=UTF8&amp;display=public&amp;page="
	# TODO fix Me This doesn't descend as expected. Cookies enabled too
	for nextPage in range(2, 3):
		print href + str(nextPage) + "<< checking..."
		parsePage(href + str(nextPage), ssheet, allProducts, nextPage)
	index = 1
	print str(len(allProducts)) + " products found."
	for product in allProducts:
		getProductPageProductData(product)
		product.writeSS(ssheet, index)
		index += 1
	return
	
# Call recursively to get the other sheets.
def parsePage(reviewURL, ssheet, allProducts, fromPage=1):
	if (reviewURL == None):
		return
	print "review URL: " + reviewURL
	if (fromPage > 1):
		refererURL = reviewURL + str(fromPage -1)
		page = reviewer.query_URL(reviewURL, refererURL)
	else:
		page = reviewer.query_URL(reviewURL)
	trs = page.split('<tr>')
	print "there are " + str(len(trs)) + " trs in the page"
	for tr in trs:
		if (tr.find(" out of 5 stars") > -1): # we have a product table record. Many data points can be gotten from here.
			allProducts.append(getReviewerPageProductData(tr))
	return

# Gets what data we can from the reviewers product page
# We go here first because there are no <read more> links 
# to follow as there are on the product pages.
# param: data - <TR> from the Reviewer's product review page.
# param: ss - SpreadSheet sheet.
# return: new Product object.
def  getReviewerPageProductData(data):
	product = Product()
	# use re to get stars: alt="4.0 out of 5 stars"
	starsPos = data.find(' out of 5 stars')
	stars = data[starsPos -3: starsPos]
	#print stars + "<<<<<<<<<<"
	product.add('reviewstar', stars)
	# get the product url
	prodUrlPos = data.find('This review')
	start = data.find('href="', prodUrlPos) + len('href="')
	end   = data.find('"', start)
	prodUrl = data[start:end]
	print prodUrl + "<<<<<<<<<< prodUrl"
	product.add('producturl', prodUrl)
	# remove all the html for easier text identification.
	text = reviewer.remove_html_tags(data)
	#print text + "<<<<<<<<<<"
	textStrings = text.split('\n')
	# reDate = re.compile(r'\d{4}$') # this doesn't work for some f*!#ing reason.
	reHelp = re.compile(r'^\d{1}')
	for textString in textStrings:
		textString = textString.strip()
		if (textString == ""):
			continue;
		#print ">>>>>>>>>" + textString + "<<<<<<<<<<"
		textString = encode_utf8(textString)
		if (reHelp.match(textString)):
			votesHelp = textString.split(' people')[0]
			votes = votesHelp.split(' of ')[1]
			helpful = votesHelp.split(' of ')[0]
			if (votes != ""):
				product.add('votes', votes)
			if (helpful != ""):
				product.add('helpful', helpful)
			print votes + " : " + helpful + " votes : helpful <<<<<<<<<<"
		if (textString.find('This review') > -1):
			title = textString.split(':')[1]
			product.add('product', title.lstrip())
			#print title + "<<<<<<<<<< title"
			# get the parens 'cause they include the catagory:
			rawCatagory = title.split('(')
			catIndex = 1
			rawCatagory.reverse()
			rawCatagory = rawCatagory[:-1]
			#print str(rawCatagory) + "<<<<<<<<+< catagory"
			for catagory in rawCatagory:
				catagory = catagory.replace(')', '')
				product.add('category'+str(catIndex), catagory.rstrip()) # trim closing paren
				catIndex += 1
		if (len(textString.split(', ')) == 3):
			title = textString.split(', ')[0]
			date  = textString.split(', ')[1]
			year  = textString.split(', ')[2]
			product.add('reviewtitle', title)
			product.add('reviewdate', date + ", " + year)
			product.add('reviewyear', year)
			#print title + " : " + date + ", " + year + "<<<<<<<<<<"
		elif (len(textString) > 120): # comments are long. gawd is that lame
			product.add('content', textString)
			product.add('characters', len(textString))
			#print str(len(textString)) + " characters long.<<<<<<<<<<"
	return product
	
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

	
# Goes to the product page and gets the canonical data about the product.
def getProductPageProductData(product):
	# if we failed to get product URL return
	url = product.get('producturl')
	if (url == None):
		return
	data = reviewer.query_URL(url)
	starsPos = data.find(' out of 5 stars')
	stars = data[starsPos -3: starsPos]
	#print stars + "<========="
	product.add('avreview', stars)
	# product introduction date seems to follow (at least) these two types:
	dateStart = data.find('first available')
	if (dateStart > -1):
		dateEnd = data.index('\n', dateStart)
		introDate = reviewer.remove_html_tags(data[dateStart:dateEnd])
		introDate = introDate.split(': ')[1]
		print introDate + "<========="
		product.add('productfirst', introDate)
	else:
		dateStart = data.find('Publication Date')
		if (dateStart > -1):
			dateEnd = data.index('\n', dateStart)
			introDate = reviewer.remove_html_tags(data[dateStart:dateEnd])
			introDate = introDate.split(': ')[1].split(' |')[0]
			#print introDate + "<========="
			product.add('productfirst', introDate)
	totalReviewsPos = data.find('customer reviews</a>)')
	if (totalReviewsPos > -1):
		totalReviews = reviewer.remove_html_tags(data[totalReviewsPos -10: totalReviewsPos]).split('>')[1]
		print totalReviews + "<========= Total reviews"
		product.add('totalrev', totalReviews)
	return

if __name__ == "__main__":
	import doctest
	doctest.testmod()
	print "You should be running reviewer.py instead."

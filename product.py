#!/usr/bin/env python
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
		self.att['votes'] = 0
		self.att['helpful'] = 0
	
	def getHeadings(self):
		return self.productHeadings
		
	def add(self, name, value):
		self.att[name] = value
		
	def toStr(self):
		if (reviewer.DEBUG):
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
def getProductReviews(reviewer_name, reviewURL, ssheet):
	writeProductHeadings(ssheet)
	print "review URL: " + reviewURL
	page = reviewer.query_URL(reviewURL)
	trs = page.split('<tr>')
	print "there are " + str(len(trs)) + " trs in the page"
	# here we will get the information we can from the Reviewer's product page:
	allProducts = []
	for tr in trs:
		if (tr.find(" out of 5 stars") > -1): # we have a product table record.
			#nextPage = getNextPageURL(page)
			# TODO add another arg: nextPage to the function below to kick off recursion.
			allProducts.append(getReviewerPageProductData(tr))
	print str(len(allProducts)) + " products found."
	# to get here we have scraped all the data from the reviewer's product page
	# now get the product data from the Products page.
	index = 1
	for product in allProducts:
	#	getProductPageProductData(product, tr)
		product.writeSS(ssheet, index)
		index += 1
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
	prodUrl      = data[start:end]
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
		if (reHelp.match(textString)):
			votesHelp = textString.split(' people')[0]
			votes = votesHelp.split(' of ')[1]
			helpful = votesHelp.split(' of ')[0]
			product.add('votes', votes)
			product.add('helpful', helpful)
			#print votes + " : " + helpful + " votes : helpful <<<<<<<<<<"
		if (textString.find('This review') > -1):
			title = textString.split(':')[1]
			product.add('product', title.lstrip())
			#print title + "<<<<<<<<<< title"
		if (len(textString.split(', ')) == 3):
			title = textString.split(', ')[0]
			date  = textString.split(', ')[1]
			year  = textString.split(', ')[2]
			product.add('reviewtitle', title)
			product.add('reviewdate', date + ", " + year)
			product.add('reviewyear', year)
			#print title + " : " + date + ", " + year + "<<<<<<<<<<"
		elif (len(textString) > 120): # comments are long.
			product.add('content', textString)
			product.add('characters', len(textString))
			#print str(len(textString)) + " characters long.<<<<<<<<<<"
	return product

	
#def getProductPageProductData(product, data):
#	return

if __name__ == "__main__":
	import doctest
	doctest.testmod()
	print "You should be running reviewer.py instead."

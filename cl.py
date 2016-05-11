### GET IMAGE SCRAPER TO WORK
### SET UP DATA TABLE FOR POWER POINT

from pptx import Presentation #lets us create presentation
from pptx.util import Inches, Pt #lets us get pptx fonts, measurements etc
import time #lets us pause for a certain time
import requests # lets us download image from url
import os # lets us execute console commands
from datetime import datetime  
from selenium import webdriver #lets us search the web
from selenium.webdriver.common.by import By #lets us search for elements
from selenium.webdriver.common.keys import Keys #lets us enter keys into webdriver
from selenium.webdriver.remote.webelement import WebElement 

def find_loan (tuition, rate):
	loan_rates = {}
	owe = (tuition * rate)
#	loan[0] = owe
	for year in range(1,16):
		loan = owe * year
		full_price = (owe*year) + tuition
		loan_rates[year] = full_price
	return loan_rates

class powerpoint (object):

	def setup (self): # Sets up front page
		title_page_layout = prs.slide_layouts[0] #Creates
		title_slide = prs.slides.add_slide(title_page_layout)
		title_page_title = title_slide.shapes.title
		title_page_subtitle = title_slide.placeholders[1]

		title_page_title.text = 'College Loan Project'
		title_page_subtitle.text = 'Programmed by Heyaw Meteke'
		
	def export (self,file_name): # Saves and exports file to folder
		file_name = 'Presentation/' + file_name + '.pptx'
		prs.save(file_name)

	def add_image_page(self):
		image_path = start.collegeImageDir
		image_page_layout = prs.slide_layouts[6]
		image_slide =prs.slides.add_slide(image_page_layout)
		
		top = Inches(1)
		height = Inches(5.5)
		left = Inches(2)
	
		img = image_slide.shapes.add_picture(image_path, left, top, height)
	def add_info_page (self, name, location, desc):
		info_page_layout = prs.slide_layouts[1]
		info_slide = prs.slides.add_slide(info_page_layout)
		modules = info_slide.shapes
		
		title_info_page = modules.title
		body_info_page = modules.placeholders[1]
		
		title_info_page.text = name
		
		textbox = body_info_page.text_frame
		textbox.text = 'Location: ' + location
		
		p = textbox.add_paragraph()
		p.text = desc
		p.font.size = Pt(15)
	
	def add_cost_page (self, loan_table, tuition):
		table_page_layout = prs.slide_layouts[5]
		table_slide = prs.slides.add_slide(table_page_layout)
		modules = table_slide.shapes
		
		modules.title.text = 'Simple Interest'
		
		rows = 3
		cols = 3
		left = top = Inches(2.0)
		width = Inches(6.0)
		height = Inches(0.8)

		table = modules.add_table(rows, cols, left, top, width, height).table
		
		table.cell(0, 0).text = 'Year' 
		table.cell(0, 1).text = 'Interest'
		table.cell(0, 2).text = 'Balance'
		for x in range(1,3):

			table.cell(x, 0).text = str(x)
			table.cell(x, 1).text = str(34)
			table.cell(x, 2).text = str(loan_table[x])	
			
class college_scrapper (object):
	collegeName = ''
	collegeImageDir = ''
	collegeLocation = ''
	collegeTuition = 0
	collegeDesc = ''' '''
	loan_table = {}
	url = ''
	def __init__ (self):
		os.system('clear')
		test = 'sdasd'
		test.center(40)
		print '\t\t\tCollege Loan Algebra Project'
		print '\t\t\tProgrammed by Heyaw Meteke'	
		driver.set_window_size(1440, 900)
		driver.get('https://www.google.com/')	

        def get_image (self):
		image_url = ''
                driver.get('https://www.yandex.com/images/')
                search_image = driver.find_element_by_name('text')
                search_image.clear()
		search_image.send_keys('logo ' + self.collegeName)
		search_image.send_keys(Keys.RETURN)

	
	#	try:
		time.sleep(4)
		print 'Searching for Image...'
		temp_image_url = driver.find_element_by_css_selector("*[class^='serp-item__link']").get_attribute('href')
		driver.get(temp_image_url)
		
		try:
			#Makes sure to close pop up if its open
			driver.find_element_by_css_selector("*[class^='popup__content']").click()
		except:
			#if no pop up, no worry - just to catch error
			pass
		try:
	
			image_url = driver.find_element_by_css_selector("*[class^='button2 button2_theme_action button2_size_m button2_type_link button2_pin_brick-clear button2_width_max sizes__download i-bem button2_js_inited']").get_attribute('href')
	
		except:
			print "First method didnt work, going to second one"
			image_url = driver.find_element_by_xpath("/html/body/div[5]/div[3]/div[2]/div[1]/div[2]/div[1]/a").get_attribute('href')	
		#except:
		#	print 'Taking longer than usual...  Slow internet?'
		#	time.sleep(2)
		#	image_url = driver.find_element_by_css_selector("*[class^='irc_but_t']").text

		
		print 'Downloading Image...'
		driver.get(image_url)
		try:
			self.collegeImageDir =  'Presentation/image.jpg'
			photo = open(self.collegeImageDir, 'wb')
			photo.write(requests.get(image_url).content)
			photo.close()
			print 'Download Completed!'

		except:
			print 'Error downloading image'

#		download_command = download_command + 
#		os.system(download_command)
		
	def college_search (self):

		college_input = raw_input('College: ')
		college_input = college_input + ' college data'
		
		search_college = driver.find_element_by_name('q')
		search_college.clear()
		search_college.send_keys(college_input)
		search_college.send_keys(Keys.RETURN)
		print 'Searching..'

		#for tries in range(1, 10):
		#try:	
		time.sleep(5)
		try:
			driver.find_element_by_partial_link_text('CollegeData').click()
		except:
			print 'Taking longer than usual... Slow internet?'
			
			driver.get('www.google.com')
			search_college = driver.find_element_by_name('q')
        	        search_college.clear()
                	search_college.send_keys(college_input)
                	search_college.send_keys(Keys.RETURN)
			time.sleep(3)

			driver.find_element_by_partial_link_text('CollegeData').click()
			#driver.find_element_by_xpath('//a[starts-with(@href, "/url?")]').click()
		try:
			print  driver.find_element_by_css_selector("*[class^='pagetitle']").text
		except:
			print 'Taking longer than usual... Slow internet?'
			time.sleep(5)
			print  driver.find_element_by_css_selector("*[class^='pagetitle']").text
		finally:
			print 'Successfully found information'

		driver.save_screenshot('screenshot.png')		
		
		############### FIND COLLEGE INFORMATION #########################
		self.collegeLocation = driver.find_element_by_xpath("//*[@id='collprofile']/div[6]/div[4]/div[2]/div[1]/p").text # LOCATION OF COLLEGE
		try:
			self.collegeName = driver.find_element_by_xpath("//*[@id='collprofile']/div[6]/div[4]/div[2]/div[1]/h1").text	# NAME OF COLLEGE
		except:
			self.collegeName = driver.find_element_by_css_selector("*[class^='citystate']").text()
		
		self.collegeTuition = driver.find_element_by_xpath("//*[@id='section1']/table/tbody/tr[2]/td").text # TUITION COST FOR COLLEGE

		try:
			self.collegeDesc = driver.find_element_by_xpath("//*[@id='cont_overview']/p").text #BRIEF DESCRIPTION OF COLLEGE
		except: 
			self.collegeDesc = driver.find_element_by_css_selector("*[class^='overviewtext']").text
		##################################################################

		######### SEPERATE DATA AND SEND TUITION TO FIND_LOAN FUNCTION #####################
		self.collegeTuition = self.collegeTuition.split('$')
		tuition = ''
		tuitionTemp = self.collegeTuition[1]
		tuitionTemp = tuitionTemp.replace(' ', "")
		for char in tuitionTemp:
			try:
				works = int(char)
				tuition+= char
			except:
				pass
		self.collegeTuition = int(tuition)
		self.loan_table = find_loan(self.collegeTuition, .04)
		#print self.loan_table


prs = Presentation()
powerpoint().setup()
#prs.save('test.pptx')

driver = webdriver.Firefox()
start = college_scrapper()
start.college_search()
start.get_image()
powerpoint().add_image_page()
powerpoint().add_info_page(start.collegeName, start.collegeLocation, start.collegeDesc)
powerpoint().add_cost_page(start.loan_table, start.collegeTuition)
powerpoint().export(start.collegeName)

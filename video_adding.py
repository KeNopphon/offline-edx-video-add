#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os, tarfile, shutil, xlrd,xlwt, datetime
import string
from lxml import etree
from six.moves import html_parser

"""
	Const declaration
	this consts will help us in case of modifications in the excel sheet
"""

"""
1.1) export excel into a list of dictionary ---->      (done)
1.2) extract course into a list of dictionary ---->    (done)
2) check the existance of course structure
	--> video does not exist, do either one of following 
		--> section does not exist, create this structure down to video
		--> subsection does not exist, create this structure down to video
		--> unit does not exist, create this structure down to video
		--> section,subsection,unit do exist, create link in unit and video 
	--> video exist, contin


"""



"""
	sheet->Video
"""
VIDEOSHEET = "video"
VIDEOINDEX = 0
VIDEOSECTION = 1
VIDEOSUBSECTION = 2
VIDEOUNIT = 3
VIDEOURL = 4
VIDEONAME = 5


COURSEPATH = "course"
EXCELNAME = "video_list.xlsm"
wb = xlrd.open_workbook(EXCELNAME)
#export_path = ""
"""
hardcoded xlsmpath must change to a parameter
"""



class Course_extraction:

	def __init__(self):
		
		self.section_path = os.path.join(COURSEPATH,'chapter')
		self.subsection_path = os.path.join(COURSEPATH,'sequential')
		self.unit_path = os.path.join(COURSEPATH,'vertical')
		self.video_path = os.path.join(COURSEPATH,'video')


	def sections(self):
		sections_files = os.listdir(self.section_path)
		self.all_section = []
		for section_file in sections_files:
			tree = etree.parse(os.path.join(self.section_path,section_file))
			root = tree.getroot()
			section_name = root.get('display_name')
			section_link = string.replace(section_file, '.xml', '')
			subsection_objs = root.findall(".sequential")
			subsection_url = []
			for subsection_obj in subsection_objs:	
				subsection_url.append(subsection_obj.get('url_name'))

			#print "section: " + str(block_name)
			#print subblock_url
			self.all_section.append({'section_link':section_link,
				'section_name':section_name,
				'assoc_subsection_url':subsection_url})
		return(self.all_section)

	def subsections(self):
		subsections_files = os.listdir(self.subsection_path)
		self.all_subsection = []
		for subsection_file in subsections_files:
			tree = etree.parse(os.path.join(self.subsection_path,subsection_file))
			root = tree.getroot()
			subsection_name = root.get('display_name')
			subsection_link = string.replace(subsection_file, '.xml', '')
			unit_objs = root.findall(".vertical")
			unit_url = []
			for unit_obj in unit_objs:	
				unit_url.append(unit_obj.get('url_name'))

			self.all_subsection.append({'subsection_link':subsection_link,
				'subsection_name':subsection_name,
				'assoc_unit_url':unit_url})
		return(self.all_subsection)


	def units(self):
		units_files = os.listdir(self.unit_path)
		self.all_unit = []
		for unit_file in units_files:
			tree = etree.parse(os.path.join(self.unit_path,unit_file))
			root = tree.getroot()
			unit_name = root.get('display_name')
			unit_link = string.replace(unit_file, '.xml', '')
			
			video_objs = root.findall(".video")
			video_url = []
			for video_obj in video_objs:	
				video_url.append(video_obj.get('url_name'))
			"""
			html_objs = root.findall(".html")
			problem_objs = root.findall(".problem")
			html_url = []
			problem_url = []
			for html_obj in html_objs:	
				unit_url.append(unit_obj.get('url_name'))
			for problem_obj in problem_objs:	
				unit_url.append(unit_obj.get('url_name'))
			"""
			self.all_unit.append({'unit_link':unit_link,'unit_name':unit_name,'assoc_video_url':video_url})
		return(self.all_unit)

	def videos(self):
		videos_files = os.listdir(self.video_path)
		self.all_video = []
		if os.path.exists(self.video_path):
			for video_file in videos_files:
				tree = etree.parse(os.path.join(self.video_path,video_file))
				root = tree.getroot()
				video_name = root.get('display_name')
				youtube_id_0_75 = root.get('youtube_id_0_75')
				youtube_id_1_0 = root.get('youtube_id_1_0')
				youtube_id_1_25 = root.get('youtube_id_1_25')
				youtube_id_1_5 = root.get('youtube_id_1_5')
				subtitle = root.get('sub')
				html5_sources = root.get('html5_sources')
				download_video = root.get('download_video')
				download_track = root.get('download_track')
				display_name = root.get('display_name')
				url_name = root.get('url_name')
				youtube = root.get('youtube')
					
				self.all_video.append({'video_name':display_name,
					'youtube_id_0_75':youtube_id_0_75,
					'youtube_id_1_0':youtube_id_1_0,
					'youtube_id_1_25':youtube_id_1_25,
					'youtube_id_1_5':youtube_id_1_5,
					'subtitle':subtitle,
					'html5_sources':html5_sources,
					'download_video':download_video,
					'download_track':download_track,
					'display_name':display_name,
					'url_name':url_name,
					'youtube':youtube,})
		else:
			os.makedirs(self.video_path) # create video folder if not exist
		return(self.all_video)



def clean_filename(s, minimal_change=False):
    """
    Sanitize a string to be used as a filename.
    If minimal_change is set to true, then we only strip the bare minimum of
    characters that are problematic for filesystems (namely, ':', '/' and
    '\x00', '\n').
    """

    # First, deal with URL encoded strings
    h = html_parser.HTMLParser()
    s = h.unescape(s)

    # strip paren portions which contain trailing time length (...)
    s = (
        s.replace(':', '-')
        .replace('/', '-')
        .replace('\x00', '-')
        .replace('\n', '')
    )

    if minimal_change:
        return s

    s = s.replace('(', '').replace(')', '')
    s = s.rstrip('.')  # Remove excess of trailing dots

    s = s.strip().replace(' ', '_')
    valid_chars = '-_.()%s%s' % (string.ascii_letters, string.digits)
    return ''.join(c for c in s if c in valid_chars)



def excel2list():
	sheetstruc = wb.sheet_by_name(VIDEOSHEET)
	all_video = []
	for row in range(1, sheetstruc.nrows):

		video_idx = sheetstruc.cell_value(row,VIDEOINDEX)
		video_section = sheetstruc.cell_value(row,VIDEOSECTION)
		video_subsection = sheetstruc.cell_value(row,VIDEOSUBSECTION)
		video_unit = sheetstruc.cell_value(row,VIDEOUNIT)
		video_url = sheetstruc.cell_value(row,VIDEOURL)
		video_url_id = video_url.rsplit('https://youtu.be/', 1)[1]
		video_name = sheetstruc.cell_value(row,VIDEONAME)


		all_video.append({'idx':video_idx,
			'section':video_section,
			'subsection':video_subsection,
			'unit':video_unit,
			'video_link':video_url,
			'video_id':video_url_id,
			'video_name':video_name}) 
	return(all_video)


	




def check_video_existance(videos_excel,course_structure):

	videos_info = []
	for video_info in course_structure.videos():
		videos_info.append(video_info['display_name'])

	for video_excel in videos_excel:
		if video_excel['video_name'] not in videos_info:
			check_section_existance(video_excel,course_structure)
			check_subsection_existance(video_excel,course_structure)
			check_unit_existance(video_excel,course_structure)


def check_section_existance(video_excel,course_structure):
	sections = course_structure.sections()
	sections_name = []
	for section in sections:
		sections_name.append(section['section_name'])

	if video_excel['section'] not in sections_name:
		print 'section: ' + video_excel['section'] + ' does not exist'
		print 'creating a new section: ' + video_excel['section'] 
		xml_creation(video_excel['section'],video_excel['subsection'],'chapter','sequential')
	else:
		print 'section: ' + video_excel['section'] + ' does exists'
		#xml_update_subelement(video_excel['section'],video_excel['subsection'],sections,'section','chapter','subsection','sequential')

def check_subsection_existance(video_excel,course_structure):
	sections = course_structure.sections()
	subsections = course_structure.subsections()
	subsections_name = []
	for subsection in subsections:
		subsections_name.append(subsection['subsection_name'])
	if video_excel['subsection'] not in subsections_name:
		print 'subsection: ' + video_excel['subsection'] + ' does not exist'
		print 'creating a new subsection: ' + video_excel['subsection'] 
		xml_creation(video_excel['subsection'],video_excel['unit'],'sequential','vertical')
		update_section_as_upper_layer(sections, video_excel)
		#xml_creation(video_excel['section'],video_excel['subsection'],'chapter','sequential')
		#xml_creation(video_excel['subsection'],video_excel['unit'],'sequential','vertical')
		#update_section_as_upper_layer(sections, video_excel)
	else:
		print 'subsection: ' + video_excel['subsection'] + ' does exists'
		xml_update_subelement(video_excel['section'],video_excel['subsection'],video_excel['unit'],sections,subsections,'section','chapter','subsection','sequential','unit','vertical')
		
		#print 'creating a new subsection: ' + video_excel['subsection'] 
	

	
		

		

def check_unit_existance(video_excel,course_structure):
	sections = course_structure.sections()
	subsections = course_structure.subsections()
	units = course_structure.units()
	
	units_name = []
	for unit in units:
		units_name.append(unit['unit_name'])
	if video_excel['unit'] not in units_name:
		print 'unit: ' + video_excel['unit'] + ' doot exist'
		print 'creating a new unit: ' + video_excel['unit']
		xml_creation(video_excel['unit'],video_excel['video_name'],'vertical','video')
		update_subsection_as_upper_layer(sections,subsections, video_excel)
 
	else:
		print 'unit: ' + video_excel['unit'] + ' does exist'
		xml_update_subelement(video_excel['subsection'],video_excel['unit'],video_excel['video_name'],subsections,units,'subsection','sequential','unit','vertical','video','video')
		
		#xml_creation(video_excel['subsection'],video_excel['unit'],'sequential','vertical')

'''
def xml_creation(layer, sub_layer, layer_name, sublayer_name):
	layer_filename      = clean_filename(layer)          
	sublayer_filename   = clean_filename(sub_layer)
	layer_file = os.path.join(COURSEPATH, layer_name, layer_filename) + ".xml";
	if os.path.exists(layer_file):   # if xml file exists
		tree = etree.parse(layer_file)
		page = tree.getroot()
		doc = etree.ElementTree(page)	
		subelements_objs = page.findall("."+sublayer_name)
		subelements = []
		for subelements_obj in subelements_objs:
			subelements.append(subelements_obj.get('url_name'))
		if sublayer_filename not in subelements: # if sub_layer element does not exist
			etree.SubElement(page, sublayer_name,url_name=sublayer_filename)
			doc.write(layer_file, pretty_print=True, xml_declaration=False, encoding='utf-8')
	else:   # if xml file does not exist, create a new xml file
		page = etree.Element(layer_name, display_name=layer) 
		doc = etree.ElementTree(page)	
		etree.SubElement(page, sublayer_name,url_name=sublayer_filename)
		doc.write(layer_file, pretty_print=True, xml_declaration=False, encoding='utf-8')
'''	
def xml_creation(layer, sub_layer, layer_name, sublayer_name):
	layer_filename      = clean_filename(layer)          
	sublayer_filename   = clean_filename(sub_layer)
	layer_file = os.path.join(COURSEPATH, layer_name, layer_filename) + ".xml";
	page = etree.Element(layer_name, display_name=layer) 
	doc = etree.ElementTree(page)		
	etree.SubElement(page, sublayer_name,url_name=sublayer_filename)
	doc.write(layer_file, pretty_print=True, xml_declaration=False, encoding='utf-8')

def xml_update_subelement(upper_layer,layer,sublayer, course_upperlayer,course_layer , edx_upperlayer_name,olx_upperlayer_name,edx_layer_name,olx_layer_name,edx_sublayer_name, olx_sublayer_name):

	sublayer_filename      = clean_filename(sublayer)          
	''' search section in course that matches those in excel file'''
	print course_upperlayer
	for e_course_upperlayer in course_upperlayer:
		if layer == e_course_upperlayer[edx_upperlayer_name+'_name']:
			match_upper = e_course_upperlayer
		else:
			return()
	
	for element in course_layer:
		if layer == element[edx_layer_name+'_name'] and element[edx_layer_name+"_link"] in match_upper['assoc_'+edx_layer_name+'_url']:
			layer_file = os.path.join(COURSEPATH, olx_layer_name, element[edx_layer_name+"_link"]) + ".xml";
			tree = etree.parse(layer_file)
			page = tree.getroot()
			doc = etree.ElementTree(page)	
			subelements_objs = page.findall("."+olx_sublayer_name)
			subelements = []
			for subelements_obj in subelements_objs:
				subelements.append(subelements_obj.get('url_name'))
			if sublayer_filename not in subelements: # if sub_layer element does not exist
				etree.SubElement(page, olx_sublayer_name,url_name=sublayer_filename)
				doc.write(layer_file, pretty_print=True, xml_declaration=False, encoding='utf-8')


def update_section_as_upper_layer(all_course_section,video_excel_info):


	section_name = video_excel_info['section']
	subsection_name =  video_excel_info['subsection']
	subsection_filename      = clean_filename(subsection_name)   
	for course_section in all_course_section:
		if section_name == course_section['section_name']:
			section_file = os.path.join(COURSEPATH,'chapter',course_section['section_link']) + ".xml";
			tree = etree.parse(section_file)
			page = tree.getroot()
			doc = etree.ElementTree(page)	
			subsection_links = []
			subsection_objs = page.findall(".sequential")
			for subsection_obj in subsection_objs:
				subsection_links.append(subsection_obj.get('url_name'))
			print subsection_links
			if subsection_filename not in subsection_links: # if sub_layer element does not exist
				etree.SubElement(page, 'sequencial',url_name=subsection_filename)
				doc.write(section_file, pretty_print=True, xml_declaration=False, encoding='utf-8')

def update_subsection_as_upper_layer(all_course_section,all_course_subsection,video_excel_info):

	section_name = video_excel_info['section']
	subsection_name =  video_excel_info['subsection']
	unit_name =   video_excel_info['unit']
	unit_filename      = clean_filename(unit_name)   
	''' search section in course that matches those in excel file'''
	for course_section in all_course_section:
		if section_name == course_section['section_name']:
			match_section = course_section
	''' search subsection in course that matches those in excel file'''
	for course_subsection in all_course_subsection:
		''' if both section and subsection are the same as those in excel file, add unit in subsection xml'''
		if subsection_name == course_subsection['subsection_name'] and course_subsection['subsection_link'] in match_section['assoc_subsection_url']:

				subsection_file = os.path.join(COURSEPATH,'sequential',course_subsection['subsection_link'])+".xml";
				tree = etree.parse(subsection_file)
				page = tree.getroot()
				doc = etree.ElementTree(page)	
				unit_links = []
				unit_objs = page.findall(".vertical")
				for unit_obj in unit_objs:
					unit_links.append(unit_obj.get('url_name'))
				if unit_filename not in unit_links: # if sub_layer element does not exist
					etree.SubElement(page, 'vertical',url_name=unit_filename)
					doc.write(subsection_file, pretty_print=True, xml_declaration=False, encoding='utf-8')

def add_unit_link(videos_excel,course_structure):

	units = course_structure.units()
	
	units = []
	for unit_info in course_structure.units():
		units.append(unit_info)

	for video_excel in videos_excel:
		for unit in units:
			if video_excel['unit'] == unit['unit_name']:
				unit_filename = os.path.join(COURSEPATH,'vertical',unit['unit_link'])+".xml";
				tree = etree.parse(unit_filename)
				page = tree.getroot()
				doc = etree.ElementTree(page)	
				video_links = []
				video_objs = page.findall(".video")

				for video_obj in video_objs:
					video_links.append(video_obj.get('url_name'))
				if video_excel['video_name'] not in video_links: # if sub_layer element does not exist
					etree.SubElement(page, 'video',url_name=video_excel['video_name'])
					doc.write(unit_filename, pretty_print=True, xml_declaration=False, encoding='utf-8')
					print 'added link to video component at unit :'+unit_filename



	


def video_component(videos):
	print "--------------------------------------- begin creating video xml file ----------------------------------------"
	for video in videos:
	
		youtube = "1.00:"+ video['video_id'].rstrip()
		youtube_id_1_0 = youtube
		display_name = video['video_name']
		urlname = video['video_name']
		download_TF = "false"
		edx_video_id = ""
		url_source = "[]"
		link_sub = video['video_id']

		video_path = os.path.join(COURSEPATH,'video',video['video_name'])
		xmlfile = video_path + ".xml";
		page = etree.Element('video', youtube=youtube, url_name = urlname, display_name = display_name, download_video=download_TF,edx_video_id =edx_video_id, html5_sources=url_source,sub=link_sub, youtube_id_1_0=youtube_id_1_0) 
		doc = etree.ElementTree(page)
		doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')
		print "video component: " + xmlfile + " is created"
	print "-------------------------------------------------------------------------------------------------------"+ "\n"


def make_tarfile():
	"""
	Packs all in a targz file ready to import.
	"""
	with tarfile.open(path + '/' + path + '.tar.gz', 'w:gz') as tar:
		for f in os.listdir(path):
			tar.add(path + "/" + f, arcname=os.path.basename(f))
		tar.close()
	print "uploadable file is created at " + path + '/' + path + '.tar.gz'




def main():

	'''
	Main script makes the calls in order to clean the resulting thir and after that generate that dir and the targz
	that we will use to import the course
	'''
	
	course_structure = Course_extraction()
	all_video = excel2list()
	#check_video_existance(all_video,course_structure)
	add_unit_link(all_video,course_structure)
	video_component(all_video)
	#print(course_structure.subsections())
	##create_directory_tree()
	#remove_existing_seq_video_link()
	#read_video_and_link_mapping()
	#video_component()
	
	#make_tarfile()






if __name__ == '__main__':
	try:
		main()
	except KeyboardInterrupt:
		logging.warn("\n\nCTRL-C detected, shutting down....")
		sys.exit(ExitCode.OK)
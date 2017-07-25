#!/usr/bin/env python
# -*- coding: utf-8 -*-


import os, tarfile, shutil, xlrd,xlwt, datetime,sys
import json
import string
from lxml import etree
from six.moves import html_parser


"""
	Const declaration
	this consts will help us in case of modifications in the excel sheet
"""



reload(sys)
sys.setdefaultencoding('utf-8')
print sys.getdefaultencoding()
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
TRANSCRIPTDIR = 6
ENTRANSCRIPTFILE = 7
JPTRANSCRIPTFILE = 8


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
		transcript_dir = sheetstruc.cell_value(row,TRANSCRIPTDIR)
		en_transcript_file = sheetstruc.cell_value(row,ENTRANSCRIPTFILE)
		jp_transcript_file = sheetstruc.cell_value(row,JPTRANSCRIPTFILE)


		all_video.append({'idx':video_idx,
			'section':video_section,
			'subsection':video_subsection,
			'unit':video_unit,
			'video_link':video_url,
			'video_id':video_url_id,
			'video_name':video_name,
			'transcript_dir':transcript_dir,
			'en_transcript_file':en_transcript_file,
			'jp_transcript_file':jp_transcript_file}) 
	return(all_video)


	


def find_video_comp_name(video_from_excel):
	video_file_ls = os.listdir(os.path.join(COURSEPATH,'video'))
	for video_file in video_file_ls:
		#print video_file
		tree = etree.parse(os.path.join(COURSEPATH,'video',video_file))
		page = tree.getroot()
		doc = etree.ElementTree(page)	
		video_objs = page.attrib['display_name']
		#print video_from_excel.encode('utf-8'),video_objs.encode('utf-8')
		if video_from_excel.encode('utf-8') == video_objs.encode('utf-8'):
			print video_from_excel
			video_path = os.path.join(COURSEPATH,'video',video_file) 
			return video_path
		#print video_objs.encode('utf-8')

def transcript2static(video_info):
	static_path = os.path.join(COURSEPATH,'static')
	transcripts = dict()
	if video_info['en_transcript_file'] != '':
		transcript_path = os.path.join(video_info['transcript_dir'],video_info['en_transcript_file'])
		shutil.copy(transcript_path, static_path)
		transcripts['en'] = str(video_info['en_transcript_file'])

	if video_info['jp_transcript_file'] != '':
		transcript_path = os.path.join(video_info['transcript_dir'],video_info['jp_transcript_file'])
		shutil.copy(transcript_path, static_path)
		transcripts['ja'] = str(video_info['jp_transcript_file'])




	return transcripts
		

############################ for creating video compoenent ##############################################

def search_excel_in_course(from_excel,course_structure):
	for row_ in from_excel:
		selected_section = find_section_name(row_,course_structure.sections())
		selected_subsection = find_subsection_name(row_,course_structure.subsections(),selected_section)
		selected_unit = find_unit_name(row_,course_structure.units(),selected_subsection)
		modify_video(row_,course_structure.videos(),selected_unit)

def find_section_name(row_section,course_section):
	
	for course_sec_row in course_section:
		course_sec_row['section_name'] = course_sec_row['section_name'].rstrip()
		row_section['section'] = row_section['section'].rstrip()
		if course_sec_row['section_name']== row_section['section']:
			print 'found section: ' + (row_section['section'])+ ' in the exported course'
			selected_section = course_sec_row
			return selected_section

	print 'no section: ' + (row_section['section']) + ' in the exported course'
	print 'create a new no section: ' + (row_section['section']) + ' in the exported course'
	new_section_link =  os.urandom(16).encode('hex')
	new_section_file =  new_section_link + '.xml'
	new_section_path =  os.path.join(COURSEPATH, 'chapter', new_section_file)
	new_subsection_link = os.urandom(16).encode('hex')
	page = etree.Element('chapter', display_name= row_section['section']) 
	doc = etree.ElementTree(page)
	etree.SubElement(page, 'sequential',url_name=new_subsection_link)
	doc.write(new_section_path, pretty_print=True, xml_declaration=False, encoding='utf-8')
	selected_section = {'section_link':new_section_link,'section_name':row_section['section'],'assoc_subsection_url':[new_subsection_link]}
	return selected_section

def find_subsection_name(row_subsection,course_subsection,selected_section):
	
	for course_subsec_row in course_subsection:
		course_subsec_row['subsection_name'] = course_subsec_row['subsection_name'].rstrip()
		row_subsection['subsection'] = row_subsection['subsection'].rstrip()
		if course_subsec_row['subsection_name']== row_subsection['subsection']:
			if course_subsec_row['subsection_link'] in selected_section['assoc_subsection_url']:
				print 'found subsection: ' + (row_subsection['subsection'])+ ' in the exported course'
				selected_subsection = course_subsec_row
				return selected_subsection

	print 'no subsection: ' + (row_subsection['subsection']) + ' in the exported course'
	print 'create a new subsection: ' + (row_subsection['subsection']) + ' in the exported course'
	new_subsection_link =  selected_section['assoc_subsection_url'][0]
	new_subsection_file =  new_subsection_link + '.xml'
	new_subsection_path =  os.path.join(COURSEPATH, 'sequential', new_subsection_file)
	new_unit_link = os.urandom(16).encode('hex')
	page = etree.Element('sequential', display_name= row_subsection['subsection']) 
	doc = etree.ElementTree(page)
	etree.SubElement(page, 'vertical',url_name=new_unit_link)
	doc.write(new_subsection_path, pretty_print=True, xml_declaration=False, encoding='utf-8')
	selected_subsection = {'subsection_link':new_subsection_link,'subsection_name':row_section['subsection'],'assoc_unit_url':[new_unit_link]}
	return selected_subsection

def find_unit_name(row_unit,course_unit,selected_subsection):
	
	for course_unit_row in course_unit:
		course_unit_row['unit_name'] = course_unit_row['unit_name'].rstrip()
		row_unit['unit'] = row_unit['unit'].rstrip()
		if course_unit_row['unit_name'] == row_unit['unit']:
			if course_unit_row['unit_link'] in selected_subsection['assoc_unit_url']:
				print 'found unit: ' + (row_unit['unit'])+ ' in the exported course'
				selected_unit = course_unit_row
				return selected_unit

	print 'no unit: ' + (row_unit['unit']) + 'in the exported course'
	print 'crate a new unit: ' + (row_unit['unit']) + ' in the exported course'
	new_unit_link =  selected_subsection['assoc_unit_url'][0]
	new_unit_file =  new_unit_link + '.xml'
	new_unit_path =  os.path.join(COURSEPATH, 'vertical', new_unit_file)
	new_video_link = os.urandom(16).encode('hex')
	page = etree.Element('vertical', display_name= row_unit['unit']) 
	doc = etree.ElementTree(page)
	etree.SubElement(page, 'video',url_name=new_video_link)
	doc.write(new_unit_path, pretty_print=True, xml_declaration=False, encoding='utf-8')
	selected_unit = {'unit_link':new_subsection_link,'unit_name':row_unit['unit'],'assoc_video_url':[new_video_link]}
	return selected_unit



############################for editing video video_component 	###############################################

def modify_video(row_video,course_video,selected_unit):
	
	
	for course_video_row in course_video:
		course_video_row['video_name'] = course_video_row['video_name'].rstrip()
		row_video['video_name'] = row_video['video_name'].rstrip()
		

		if course_video_row['video_name'] == row_video['video_name']:
			if course_video_row['url_name'] in selected_unit['assoc_video_url']:
				
				print 'found video component: ' + (row_video['video_name'])+ ' in the exported course'
				print 'remove video: ' +(row_video['video_name']) + ' from exported file'
				video_file =  course_video_row['url_name'] + '.xml'
				video_path =  os.path.join(COURSEPATH, 'video', video_file)
				os.remove(video_path)
				print 'add a new video: ' +(row_video['video_name']) + ' to exported file'
				youtube = '1.00:'+ row_video['video_id'].rstrip()
				youtube_id_1_0 = row_video['video_id'].rstrip()
				display_name = row_video['video_name']
				urlname = course_video_row['url_name']
				download_TF = "false"
				edx_video_id = ""
				url_source = "[]"
				link_sub = row_video['video_id']
				transcripts = transcript2static(row_video)
				if transcripts != []:
					page = etree.Element('video', youtube=youtube, url_name = urlname, display_name = display_name, download_video=download_TF,edx_video_id =edx_video_id, html5_sources=url_source,sub=link_sub, youtube_id_1_0=youtube_id_1_0,transcripts=str(json.dumps(transcripts))) 
					for key, value in transcripts.iteritems():
						etree.SubElement(page, 'transcript',language=key,src=value)
				else: 
					page = etree.Element('video', youtube=youtube, url_name = urlname, display_name = display_name, download_video=download_TF,edx_video_id =edx_video_id, html5_sources=url_source,sub=link_sub, youtube_id_1_0=youtube_id_1_0,transcripts="") 
				doc = etree.ElementTree(page)
				doc.write(video_path, pretty_print=True, xml_declaration=False, encoding='utf-8')
				print "video component: " + video_path + " is created"
				print "------------------------------------------------------------\n\n"
				return



	print 'no video: ' + (row_video['video_name']) + ' in the exported course'
	print 'crate a new video: ' + (row_video['video_name']) + ' in the exported course'

	video_file = selected_unit['assoc_video_url'][0] + '.xml'
	video_path =  os.path.join(COURSEPATH, 'video', video_file)
	youtube = '1.00:'+ row_video['video_id'].rstrip()
	youtube_id_1_0 = row_video['video_id'].rstrip()
	display_name = row_video['video_name']
	urlname = selected_unit['assoc_video_url'][0]
	download_TF = "false"
	edx_video_id = ""
	url_source = "[]"
	link_sub = row_video['video_id']
	transcripts = transcript2static(row_video)
	if transcripts != []:
		page = etree.Element('video', youtube=youtube, url_name = urlname, display_name = display_name, download_video=download_TF,edx_video_id =edx_video_id, html5_sources=url_source,sub=link_sub, youtube_id_1_0=youtube_id_1_0,transcripts=str(json.dumps(transcripts))) 
		for key, value in transcripts.iteritems():
			etree.SubElement(page, 'transcript',language=key,src=value)
	else: 
		page = etree.Element('video', youtube=youtube, url_name = urlname, display_name = display_name, download_video=download_TF,edx_video_id =edx_video_id, html5_sources=url_source,sub=link_sub, youtube_id_1_0=youtube_id_1_0,transcripts="") 
	doc = etree.ElementTree(page)
	doc.write(video_path, pretty_print=True, xml_declaration=False, encoding='utf-8')
	print 'add a new video: ' + row_video['video_name'] + ' to exported file'
	print "------------------------------------------------------------\n\n"

################################################################################################################


############################for editing video video_component 	###############################################

def edited_video_component(videos):
	print "--------------------------------------- begin creating video xml file ----------------------------------------"
	for video in videos:
	
		youtube = video['video_id'].rstrip()
		youtube_id_1_0 = '1.00:'+ youtube
		display_name = video['video_name']
		urlname = video['video_name']
		download_TF = "false"
		edx_video_id = ""
		url_source = "[]"
		link_sub = video['video_id']
		video_path = find_video_comp_name(display_name)
		transcripts = transcript2static(video)
		page = etree.Element('video', youtube=youtube, url_name = urlname, display_name = display_name, download_video=download_TF,edx_video_id =edx_video_id, html5_sources=url_source,sub=link_sub, youtube_id_1_0=youtube_id_1_0) 
		if transcripts != []:
			for transcript in transcripts:
				key = transcript.keys() 
				etree.SubElement(page, 'transcript',language=key[0],src=transcript[key[0]])

		doc = etree.ElementTree(page)
		doc.write(video_path, pretty_print=True, xml_declaration=False, encoding='utf-8')
		#print "video component: " + xmlfile + " is created"
		print "video component: " + video_path + " is created"
	print "-------------------------------------------------------------------------------------------------------"+ "\n"

################################################################################################################



def make_tarfile():
	"""
	Packs all in a targz file ready to import.
	
	path = COURSEPATH
	with tarfile.open(path + '/' + path + '.tar.gz', 'w:gz') as tar:
		for f in os.listdir(path):
			print f
			tar.posix
			tar.add(path + "/" + f, arcname=os.path.basename(f))
		tar.close()
	print "uploadable file is created at " + path + '/' + path + '.tar.gz'
	"""
	addpath = 'set PATH=%PATH%;C:\Program Files\7-Zip\ ' 
	compress_tar = '7z a course.tar course\ '
	compress_targz = '7z a course.tar.gz course.tar'
	os.system(addpath)
	os.system(compress_tar)
	os.system(compress_targz)
	os.remove('course.tar')




def main():

	search_excel_in_course(excel2list(),Course_extraction())
	make_tarfile()

	






if __name__ == '__main__':
	try:
		main()
	except KeyboardInterrupt:
		logging.warn("\n\nCTRL-C detected, shutting down....")
		sys.exit(ExitCode.OK)







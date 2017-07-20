# coding=utf-8

import os, tarfile, shutil, xlrd,xlwt, datetime
from lxml import etree


"""
	Const declaration
	this consts will help us in case of modifications in the excel sheet
"""

"""
1) export excel into a list of dictionary
2) check the existance of course structure
	--> video does not exist, do either one of following 
		--> section does not exist, create this structure down to video
		--> subsection does not exist, create this structure down to video
		--> unit does not exist, create this structure down to video
		--> section,subsection,unit do exist, create link in unit and video 
	--> video exist, continue


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


PATH = "course"
EXCELNAME = "video_list.xlsm"
wb = xlrd.open_workbook(EXCELNAME)
#export_path = ""
"""
hardcoded xlsmpath must change to a parameter
"""




def create_directory_tree():
	"""
	Generates the directory structure needed for our xml project
	"""
	print "\n"
	print "-------------------------------------- check directory ----------------------------------------------------"
	print os.getcwd()
	if os.path.exists(path):
		if not os.path.exists(path + "/vertical"):
			os.makedirs(path + "/vertical")
			print "++++++++++++++++++ ["+ path +"/verical]" + " folder is created ++++++++++++++++++" 
		if not os.path.exists(path + "/video"):
			os.makedirs(path + "/video")
			print "++++++++++++++++++ ["+ path +"/video]" + " folder is created ++++++++++++++++++"
		else:
			shutil.rmtree(path+ "/video", ignore_errors=False, onerror=None)
			os.makedirs(path + "/video")
			print "++++++++++++++++++ removed existing ["+ path +"/video]" + " folder and video files ++++++++++++++++++"
			print "++++++++++++++++++"+ path +"/video" + " folder is created ++++++++++++++++++" 
	else:
		print "could not find a course's directory"
		print "-------------------------------------------------------------------------------------------------------"+ "\n"
		exit(0)
	print "-------------------------------------------------------------------------------------------------------"+ "\n"
		

def remove_existing_seq_video_link():
	print "---------------------- begin removing an existing video link ----------------------------------"
	seq_path = "course/sequential"
	ver_path = "course/vertical"
	seq_ls = os.listdir(seq_path)
	for each_seq in seq_ls:
		tree = etree.parse(seq_path+"/"+each_seq)
		root = tree.getroot()
		ver_ls = root.findall(".vertical")
		if ver_ls:
			for each_ver in ver_ls:
				ver_filename = each_ver.get('url_name') + ".xml"
				ver_tree = etree.parse(ver_path+"/"+ver_filename)
				ver_root = ver_tree.getroot()
				video_comp_ls = ver_root.findall(".//video")
				comp_ls = ver_root.findall(".//")
				if video_comp_ls:
					if len(video_comp_ls) == len(comp_ls):
						os.remove(ver_path+"/"+ver_filename)
						each_ver.getparent().remove(each_ver)
						print "remove Units " + ver_root.get('display_name') + " from subsection " + root.get('display_name') + ":" + each_seq
						tree.write(seq_path+"/"+each_seq, pretty_print=True, xml_declaration=False, encoding='utf-8')
					else:
						print each_ver.tag, each_ver.attrib
						for each_video in video_comp_ls:
							each_video.getparent().remove(each_video) 
							print "remove video linkID " + each_video.get('url_name') + "from Units " + ver_root.get('display_name') + " ,subsection " + root.get('display_name')
						ver_tree.write(ver_path+"/"+ver_filename, pretty_print=True, xml_declaration=False, encoding='utf-8')

			
					#print len(video_comp_ls),(comp_ls)
	print "-------------------------------------------------------------------------------------------------------"+ "\n"



	
			


def read_video_and_link_mapping():
	print "---------------------------------begin mapping video link from excel file----------------------------"

	#xmlfile = path + "/course.xml"
	global sheetstruc
	sheetstruc = wb.sheet_by_name(VIDEOSHEET)
	for row in range(1, sheetstruc.nrows):

		video_url = sheetstruc.cell_value(row,VIDEOURL)
		video_url_id = video_url.rsplit('https://youtu.be/', 1)[1]
		video_name = sheetstruc.cell_value(row,VIDEONAME)
		print "+++++++++++++++++++" + video_url_id + ": "+ video_name+ "+++++++++++++++++++++++++"
		
		map_video_chapter(row)
	print "-------------------------------------------------------------------------------------------------------"+ "\n"
		
		
		
		

def map_video_chapter(_row):
	#sheetstruc = wb.sheet_by_name(VIDEOSHEET)
	current_chapter_name = sheetstruc.cell_value(_row,VIDEOSECTION)
	chap_path = "course/chapter"
	seq_path = "course/sequential"
	chap_ls = os.listdir(chap_path)
	for each_chap in chap_ls:
		tree = etree.parse(chap_path+"/"+each_chap)
		root = tree.getroot()
		if current_chapter_name == root.get('display_name'):
			print "Sections: " + str(current_chapter_name)
			seq_ls = root.findall(".sequential")
			map_video_seq(seq_ls,_row)

def map_video_seq(_current_seq_list,_row):
	#sheetstruc = wb.sheet_by_name(VIDEOSHEET)
	current_seq_name = sheetstruc.cell_value(_row,VIDEOSUBSECTION)
	seq_path = "course/sequential"
	seq_ls = os.listdir(seq_path)
	for each_seq in _current_seq_list:
		seq_filename = each_seq.get('url_name') + ".xml"
		tree = etree.parse(seq_path+"/"+seq_filename)
		root = tree.getroot()
		if current_seq_name == root.get('display_name'):
			print "Subsections: " + current_seq_name 
			ver_ls = root.findall(".vertical")
			map_video_ver(tree,root,seq_filename,ver_ls,_row)

def map_video_ver(_seq_tree,_seq_root,_seq_xmlfile,_current_ver_ls,_row):
	
	seq_path = "course/sequential"
	ver_path = "course/vertical"
	ver_ls = os.listdir(ver_path)
	ver_urlName = str(sheetstruc.cell_value(_row,VIDEOUNIT))
	ver_display_name = ver_urlName;
	video_urlName = str(sheetstruc.cell_value(_row,VIDEONAME))
	video_display_name = video_urlName
	
	# add subelement 'vertical' into sequential xml file
	etree.SubElement(_seq_root, 'vertical',url_name=ver_urlName)
	_seq_tree.write(seq_path+"/"+_seq_xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')

	# Create 'vertical' xml file with subelement video
	xmlfile = ver_path +"/"+ ver_display_name + ".xml";
	page = etree.Element('vertical', display_name=ver_display_name) 
	doc = etree.ElementTree(page)								    
	etree.SubElement(page, 'video',url_name=video_urlName)
	doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')
	print "Units: " + ver_display_name + "\n"

def video_component():
	print "--------------------------------------- begin creating video xml file ----------------------------------------"
	for row in range(1, sheetstruc.nrows):
		video_url = str(sheetstruc.cell_value(row,VIDEOURL))
		video_url_id = video_url.rsplit('https://youtu.be/', 1)[1]
		video_url_id = video_url_id.rstrip()
		video_name = str(sheetstruc.cell_value(row,VIDEONAME))
		youtube = "1.00:"+ video_url_id.rstrip()
		urlName = video_name
		download_TF = "false"
		edx_video_id = ""
		url_source = "[]"
		link_sub = ""

		video_path = "course/video"
		xmlfile = video_path +"/"+ video_name + ".xml";
		page = etree.Element('video', youtube=youtube, url_name = urlName, display_name = video_name, download_video=download_TF,edx_video_id =edx_video_id, html5_sources=url_source,sub=link_sub, youtube_id_1_0=video_url_id) 
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
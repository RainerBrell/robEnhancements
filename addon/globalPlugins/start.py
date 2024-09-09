# -*- coding: UTF-8 -*-
"""
 ROB enhancements for NVDA 
 This file is covered by the GNU General Public License.
 See the file COPYING for more details.
 Copyright (C) 2024 Rainer Brell nvda@brell.net 
 For file access I used code from the NAO project
 Thanks for the permission: Alessandro Albano, Davide De Carne and Simone Dal Maso

 *** history *** 
 2023.10.18 
 * NVDA+ctrl+F4 - Current URL, View, Copy
 * Detects empty folders in Windows 10 Explorer
 2024.03.08 
 * shift+alt+j: Jumps to the next folder  unread emails in Outlook
 * Detects empty folders in Outlook
 * insert installtasks.py 
 * ready for NVDA 2024.1 
 2024.09.09
 * Logged into the translation system
 * nvda+alt+space: markdown file viewer 
 * nvda+shift+alt+space: Saves markdown file as html
 
"""

import globalPluginHandler
from scriptHandler import script
import ui
from tones import beep 
import api
import os 
import sys 
import controlTypes
import scriptHandler
from .framework.storage import explorer
from .myExplorer import * 
from .myMarkdown import getHtmlText 
import addonHandler
addonHandler.initTranslation()

AddOnPath = os.path.dirname(__file__)

def getFileName():
	try:
		filename = explorer.get_selected_file()
	except:
		filename = None
	return filename

def isBrowser():
	"""
	 Verifies that NVDA is in a browser.
	""" 
	obj = api.getFocusObject()
	if obj.treeInterceptor:
		return True
	else:
		return False

def getCurrentDocumentURL():
	""" 
		Get current masked document URL 
	"""
	URL = None 
	obj = api.getFocusObject()
	try:
		URL = obj.treeInterceptor.documentConstantIdentifier
	except:
		return None 
	return URL 

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	# Translators: Name of the category for the keyboard mapping dialog
	scriptCategory = _("ROB enhancements")
	
	def __init__(self):
		#super(globalPluginHandler.GlobalPlugin, self).__init__()
		super().__init__()

	@script(
		# Translators: Shows the current URL of the document, press twice = copies to the clipboard.
		description=_("Show document URL, press twice copies to clipboard."),
		gesture="kb:NVDA+Control+f4"
	)
	def script_ShowDocumentURL(self, gesture):
		if isBrowser():
			URL = getCurrentDocumentURL() 
			if URL:
				if scriptHandler.getLastScriptRepeatCount() == 0:
					ui.message(URL) 
				elif scriptHandler.getLastScriptRepeatCount() == 1:
					api.copyToClip(URL)
					# Translators: URL copied to clipboard 
					ui.message(_("copied to clipboard {URL}.").format(URL=URL))
			else: 
				# Translators: No URL found in browser document
				ui.message(_("Document URL not found."))
		else:
			# Translators: The user is not in a browser.
			ui.message(_("No browser window found."))

	@script(
		# Translators: Shows the current markdown file in the NVDA browser
		description=_("Shows the current markdown file in the NVDA browser"),
		gesture="kb:NVDA+Alt+space"
	)
	def script_ShowMarkdown(self, gesture):
		markdownFile = getFileName()
		htmlFile     = os.environ["TEMP"] + r"\robEnhancements_markdown.html"
		if markdownFile:
			if markdownFile.split(".")[-1].lower() != "md":
				# Translators: it is not a markdown file
				ui.message(_("{markdownFile} is not a Markdown file").format(markdownFile=markdownFile))
				return 
		else:
			# Translators: No markdown file selected 
			ui.message(_("No markdown file selected"))
			return 
		with open(markdownFile, 'r', encoding='utf-8') as file:
			markdownContent = file.read()
		html = getHtmlText(markdownContent)
		with open(htmlFile, 'w', encoding='utf-8') as file:
			file.write(html)
		markdownFile = markdownFile.split("\\")[-1].lower().replace(".md", "")
		# translators: HTML viewer title
		htmlTitle = _("Content of {markdownFile} in HTML").format(markdownFile=markdownFile)
		ui.browseableMessage(html, title=htmlTitle, isHtml=True)


	@script(
		# Translators: Saves the selected markdown file as an HTML file
		description=_("Saves the selected markdown file as an HTML file"),
		gesture="kb:NVDA+shift+Alt+space"
	)
	def script_SaveMarkdownToHtml(self, gesture):
		markdownFile = getFileName()
		if markdownFile:
			if markdownFile.split(".")[-1].lower() != "md":
				# Translators: it is not a markdown file
				ui.message(_("{markdownFile} is not a Markdown file").format(markdownFile=markdownFile))
				return 
		else:
			# Translators: No markdown file selekted
			ui.message(_("No markdown file selected"))
			return 
		with open(markdownFile, 'r', encoding='utf-8') as file:
			markdownContent = file.read()
		currentPath = os.path.dirname(markdownFile)
		if not os.access(currentPath, os.W_OK):
			# Translators: No access to the folder
			ui.message(_("No access to the folder {currentPath}").format(currentPath=currentPath))
			return 
		html = getHtmlText(markdownContent)
		htmlFile = markdownFile[:-2] + "html"
		try:
			with open(htmlFile, 'w', encoding='utf-8') as file:
				file.write(html)
		except:
			# Translators: The HTML file could not be created
			ui.message(_("The HTML file {htmlFile} could not be created.").format(htmlFile=htmlFile))
			return 
		# Translators: The HTML file was written successfully
		msg = _("The HTML file {htmlFile} was written successfully").format(htmlFile=htmlFile)
		ui.message(msg) 



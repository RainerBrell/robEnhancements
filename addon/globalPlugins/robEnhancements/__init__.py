# -*- coding: UTF-8 -*-
"""
 ROB enhancements for NVDA 
 This file is covered by the GNU General Public License.
 See the file COPYING for more details.
 Copyright (C) 2024-2025 Rainer Brell nvda@brell.net 
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
 2024.09.09:
 * nvda+alt+space: markdown file viewer 
 * nvda+shift+alt+space: Saves markdown file as html
 2025.08.05:
 * delete explorer.py, no longer support for windows 10 explorer 
 * ready for NVDA 2025 
 2025.09.01:
 * nvda+shift+v: Taskname, 32/64bit, CPU usage, version, productname 
 2025.10.14:
 * little bugfix 
 
"""

import globalPluginHandler
from scriptHandler import script
from core import callLater 
import ui
from tones import beep 
import api
import os 
import sys 
import controlTypes
import psutil
import scriptHandler
from .framework.storage import explorer
from .myMarkdown import getHtmlText 
from .skipTranslation import translate
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
	
def get_cpu_usage(pid):
	try:
		process = psutil.Process(pid)
		# The measurement initializes the first query
		process.cpu_percent(interval=None)
		# Short break to enable a sensible measurement
		cpu_usage = process.cpu_percent(interval=1.0)
		return cpu_usage
	except psutil.NoSuchProcess:
		return f"no process with pid {pid}."
	except Exception as e:
		return f"error: {e}"

def get_process_name(focus):
	try:
		return focus.appModule.appName
	except Exception as e:
		return f"Error: {e}"
		
def get_64_32_bit(focus):
	try:
		if focus.appModule.is64BitProcess:
			# Translators: 64 bit process 
			return _("64bit")
		else:
			# Translators: 32 bit process 
			return _("32bit")
	except Exception as e:
			return f"Error: {e}"
			
def get_product_name(focus):
	try:
		pn = focus.appModule.productName
		if pn:
			return pn 
		else: 
			return translate("unknown")
	except Exception as e:
		return f"Error: {e}"
		

		
def get_product_version(focus):
	try:
		return focus.appModule.productVersion
	except Exception as e:
		return f"Error: {e}"
		
def copy_to_clip(msg):
	beep(400, 400)
	api.copyToClip(msg)
	# Translators: message  copied to clipboard 
	ui.message(_("Copied into the clipboard."))
	
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

	@script(
		# Translators: Name and current CPU usage of the process
		description=_("Shows the current program Name and CPU usage."),
		gesture="kb:NVDA+shift+v"
	)
	def script_show_current_name_cpu_task(self, gesture):
		focus = api.getFocusObject()
		if focus:
			# Translators: Wait 1 second to determine the CPU usage 
			ui.message(_("Please wait..."))
			pid         = focus.processID
			cpu         = get_cpu_usage(pid)
			appname     = get_process_name(focus)
			is64bit     = get_64_32_bit(focus)
			productname = get_product_name(focus)
			version     = get_product_version(focus)
			if appname == productname: 
				productname = ""
			msg: str = _(
				"{appname} ({is64bit}, {cpu}%) {version} {productname}"
			).format(
				appname=appname, 
				is64bit=is64bit, 
				cpu=cpu, 
				productname=productname, 
				version=version
			)
			if scriptHandler.getLastScriptRepeatCount() == 0:
				ui.message(msg)
			elif scriptHandler.getLastScriptRepeatCount() == 1:
				callLater(1000, lambda: copy_to_clip(msg))
		else:
			ui.message("No focus found")

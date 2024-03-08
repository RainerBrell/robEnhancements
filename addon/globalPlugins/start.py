# -*- coding: UTF-8 -*-
"""
 ROB enhancements for NVDA 
 This file is covered by the GNU General Public License.
 See the file COPYING for more details.
 Copyright (C) 2024 Rainer Brell nvda@brell.net 

 *** history *** 
 2023.10.18 
 * NVDA+ctrl+F4 - Current URL, View, Copy
 * Detects empty folders in Windows Explorer
 2024.03.08 
 * insert installtasks.py 
 + ready for NVDA 2024.1 
 """

import globalPluginHandler
from scriptHandler import script
import ui
from tones import beep 
import api
import os 
import controlTypes
import scriptHandler
from .myExplorer import * 
import addonHandler
addonHandler.initTranslation()

AddOnPath = os.path.dirname(__file__)

def isBrowser():
	"""
	 Verifies that NVDA is in a browser.
	""" 
	obj = api.getFocusObject()
	if not obj.treeInterceptor:
		return False 
	else:
		return True

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
	
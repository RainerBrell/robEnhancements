# -*- coding: UTF-8 -*-
"""
 ROB enhancements - add-on for NVDA 
 outlook sub add-on 
 This file is covered by the GNU General Public License.
 See the file COPYING for more details.
 Copyright (C) 2024 Rainer Brell nvda@brell.net 
 Mod: 2024.03.08 
"""

import appModuleHandler
from nvdaBuiltin.appModules import outlook
from scriptHandler import script
from NVDAObjects.UIA import ListItem, UIA
from NVDAObjects.IAccessible import IAccessible
import controlTypes
import textInfos 
import speech 
import config
import tones
import os 
import ui
import api 
import languageHandler
import webbrowser
from core import callLater 
import addonHandler

addonHandler.initTranslation()

def isValidVersion():
	obj = api.getFocusObject()
	appVersionMajor = int(obj.appModule.productVersion.split('.')[0])
	if appVersionMajor < 13: # outlook 2010 
		# Translators: unsupported Outlook version 
		msg = _("Your Outlook version {Versionsnumber} is not supported.").format(Versionsnumber=appVersionMajor)
		ui.message(msg)
		return False
	else:
		return True 

class AppModule(outlook.AppModule):

	# Translators: Name of the category for the keyboard mapping dialog 
	scriptCategory = _("ROB enhancements")
	
	def event_gainFocus(self, obj, nextHandler):
		if obj.role == controlTypes.Role.PANE: 
			callLater(300, lambda: self.emptyFolder(obj))
		if obj.windowClassName == "OutlookGrid" or obj.windowClassName == "SUPERGRID" and obj.role == 15 and obj.windowControlID == 4704:
			# Only for german 
			if languageHandler.getLanguage().split("_")[0].lower() == "de":
				if obj.name.startswith("Ungelesen "):
					obj.name = obj.name.replace("Ungelesen ", "Neu ", 1)
			else:
				f = obj.children[0].name
			config.conf["documentFormatting"]["reportTableHeaders"] = 0
			config.conf["documentFormatting"]["reportTableCellCoords"] = False
		nextHandler()
		
	def emptyFolder(self, obj):
		if obj.windowClassName == "rctrl_renwnd32" and obj.windowControlID == 0:
			focus = api.getFocusObject()
			if focus.role == controlTypes.Role.PANE: 
				#tones.beep(100, 100)
				# Translators: empty outlook folder 
				ui.message(_("This folder is empty."))

	@script(
		#Translators: Jumps to the next Outlook folder with unread mails below the inbox.
		description=_("Jumps to the next Outlook folder with unread mails below the inbox."),
		gesture="kb:alt+shift+j"
	)
	def script_JumpToNextFolderWithUnreadItems(self, gesture):	
		if not isValidVersion(): 
			return 
		try:
			dom = api.getFocusObject().appModule.nativeOm
			UnreadCount = 0 
			if dom: 
				folderName = dom.ActiveExplorer().CurrentFolder.Name
				myInbox = dom.GetNamespace("MAPI").GetDefaultFolder(6)
				Folders = myInbox.Folders
				FoldersCount = Folders.Count 
				for Folder  in Folders:
					UnreadCount = Folder.UnReadItemCount
					if UnreadCount > 0:
						# Translators: unread mails found in folder
						msg = _("{count} mails in the folder {folder}").format(count=UnreadCount, folder=Folder.Name)
						ui.message(msg)
						dom.ActiveExplorer().CurrentFolder = Folder 
						return
				# Translators: No unread mails found 
				ui.message(_("No unread mails found"))
			else:
				# Translators: if the outlook object modell is not available
				out = _("Outlook object  not available - please contact the addon developer")
				ui.browseableMessage(out)
		except:
			ui.message("Error in jump function")


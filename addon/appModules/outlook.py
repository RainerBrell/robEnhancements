# -*- coding: UTF-8 -*-
"""
 ROB enhancements - add-on for NVDA 
 outlook sub add-on 
 This file is covered by the GNU General Public License.
 See the file COPYING for more details.
 Copyright (C) 2024 Rainer Brell nvda@brell.net 
 last Modify: September 2024 
 - GoToFolders 
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

AddOnName = addonHandler.getCodeAddon().manifest['name']
sectionName = AddOnName

def initConfiguration():
	confspec = { 
		"Folder1": "string(default='')",
		"Folder2": "string(default='')",
		"Folder3": "string(default='')",
		"Folder4": "string(default='')",
		"Folder5": "string(default='')"
	}
	config.conf.spec[sectionName] = confspec

initConfiguration()

def getINI(key):
	"""  get nvda.ini value """ 
	value = config.conf[sectionName][key]
	return value 

def setINI(key, value):
	"""  set nvda.ini value """ 
	try:
		config.conf[sectionName][key] = value
	except:
		pass 
		
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
		
def SetFolder(nr):
	if not isValidVersion(): 
		return 
	try:
		dom = api.getFocusObject().appModule.nativeOm
		if dom:
			folderPath = dom.ActiveExplorer().CurrentFolder.FolderPath
			if   nr == 1: setINI("Folder1", folderPath)
			elif nr == 2: setINI("Folder2", folderPath)
			elif nr == 3: setINI("Folder3", folderPath)
			elif nr == 4: setINI("Folder4", folderPath)
			elif nr == 5: setINI("Folder5", folderPath)
			else: 
				ui.message("Error - unknown folder nomber")
			folderPath = folderPath.replace("\\\\", "")
			# Translators: Outlook mail Folder {nr} saved 
			msg = _("Folder {nr} saved as {folderPath}").format(nr=nr, folderPath=folderPath)
			ui.message(msg)
	except:
		# Translators: Error saving folder
		msg = _("Error saving folder")
		ui.message(msg) 

	
def GoToFolder(nr):
	if not isValidVersion(): 
		return 
	if   nr == 1: newPath = getINI("Folder1")
	elif nr == 2: newPath = getINI("Folder2")
	elif nr == 3: newPath = getINI("Folder3")
	elif nr == 4: newPath = getINI("Folder4")
	elif nr == 5: newPath = getINI("Folder5")
	else: newPath = ""
	if not newPath:
		# Translators: No folder set yet
		msg = _("The folder {nr} has not been specified yet. Please specify a folder to jump to first.").format(nr=nr)
		ui.message(msg) 
		return 
	newPathList = newPath.split("\\") 
	account     = newPathList[2]
	folderList  = newPathList[3:]
	try:
		dom = api.getFocusObject().appModule.nativeOm
		if dom:
			nameSpace = dom.GetNamespace("MAPI")
			existsAccount = False 
			for acc in nameSpace.Folders: 
				if acc.name == account:
					existsAccount = True 
					existsFolder = False 
					index = 0
					currentFolder = acc.folders
					for entry in folderList:
						index += 1
						for folder in currentFolder:
							if folder.name == entry: 
								newFolder = folder.folders 
								if index == len(folderList):
									# Translators: Go to the folder
									msg = _("Go to the folder {entry}").format(entry=entry)
									ui.message(msg) 
									dom.ActiveExplorer().CurrentFolder = folder 
									return 
						currentFolder = newFolder
					if not existsFolder:
						# Translators: Can not found the Folder 
						msg = _("Can not found the folder {folderList}").format(folderList=folderList)
						ui.message(msg) 
			if not existsAccount:
				# Translators: Can not found the outlook account 
				msg = _("Can not found the account: {account}").format(account=account)
				ui.message(msg) 
	except: 
		# Translators: Error, cannot go to the specified path.
		msg = _("Error, cannot go to the specified path {newPath}.").foramt(newPath=newPath)
		ui.message(msg) 

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
				#myInbox = dom.GetNamespace("MAPI").Folders("rainer@brell.net").Folders("Posteingang")
				Folders = myInbox.Folders
				FoldersCount = Folders.Count 
				for Folder in Folders:
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
		except COMError:
			ui.message("Error in jump function")

	@script(
		#Translators: set outlook mail folder 1
		description=_("Set outlook mail folder 1"),
		gesture="kb:alt+control+shift+i" 
	)
	def script_SetFolder1(self, gesture):	
		SetFolder(1)
	
	@script(
		#Translators: Go to outlook mail folder 1
		description=_("Go to outlook mail folder 1"),
		gesture="kb:alt+shift+i"
	)
	def script_GoToFolder1(self, gesture):	
		GoToFolder(1)

	@script(
		#Translators: set outlook mail folder 2
		description=_("Set outlook mail folder 2"),
		gesture="kb:" 
	)
	def script_SetFolder2(self, gesture):	
		SetFolder(2)
	
	@script(
		#Translators: Go to outlook mail folder 2
		description=_("Go to outlook mail folder 2"),
		gesture="kb:"
	)
	def script_GoToFolder2(self, gesture):	
		GoToFolder(2)

	@script(
		#Translators: set outlook mail folder 3
		description=_("Set outlook mail folder 3"),
		gesture="kb:" 
	)
	def script_SetFolder3(self, gesture):	
		SetFolder(3)
	
	@script(
		#Translators: Go to outlook mail folder 3
		description=_("Go to outlook mail folder 3"),
		gesture="kb:"
	)
	def script_GoToFolder3(self, gesture):	
		GoToFolder(3)

	@script(
		#Translators: set outlook mail folder 4
		description=_("Set outlook mail folder 4"),
		gesture="kb:" 
	)
	def script_SetFolder4(self, gesture):	
		SetFolder(4)
	
	@script(
		#Translators: Go to outlook mail folder 4
		description=_("Go to outlook mail folder 4"),
		gesture="kb:"
	)
	def script_GoToFolder4(self, gesture):	
		GoToFolder(4)

	@script(
		#Translators: set outlook mail folder 5
		description=_("Set outlook mail folder 5"),
		gesture="kb:" 
	)
	def script_SetFolder5(self, gesture):	
		SetFolder(5)
	
	@script(
		#Translators: Go to outlook mail folder 5
		description=_("Go to outlook mail folder 5"),
		gesture="kb:"
	)
	def script_GoToFolder5(self, gesture):	
		GoToFolder(5)

# -*- coding: UTF-8 -*-
# NVDA ROB enhancements / myExplorer 
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2024 Rainer Brell nvda@brell.net 

import globalPluginHandler
from tones import beep 
import ui
import controlTypes
from core import callLater 
import api 

class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	staticTextHistory = []

	def __init__(self, *args, **kwargs):
		#super(GlobalPlugin, self).__init__(*args, **kwargs)
		super().__init__()
		
	def terminate(self):
		super(GlobalPlugin, self).terminate()

	def event_nameChange(self, obj, nextHandler):
		self.appName = obj.appModule.appName
		if self.appName == "explorer": 
			if obj.role == controlTypes.ROLE_STATICTEXT and obj.windowClassName == "DirectUIHWND":
				focus = api.getFocusObject()
				if focus.childCount == 0 and obj.UIAAutomationId == "EmptyText":
					listItem = obj.name 
					callLater(500, lambda: self.addTo(listItem))
		nextHandler()

	def addTo(self, listItem):
		focus = api.getFocusObject()
		try: 
			name = focus.name 
			if name == "":
				# if is empty folder 
				self.staticTextHistory.append(listItem)
				lenght = (len(self.staticTextHistory) - 1)
				if lenght >= 1 and (self.staticTextHistory[lenght] == self.staticTextHistory[lenght-1]):
					# Solution for double output of static Text 
					self.staticTextHistory = []
					ui.message(listItem) 
		except:
			pass

# -*- coding: UTF-8 -*-
"""
 ROB enhancements for NVDA - Module myMarkdown 
 This file is covered by the GNU General Public License.
 See the file COPYING for more details.
 Copyright (C) 2025 Rainer Brell nvda@brell.net 
"""
 
import globalPluginHandler
import sys 
import os 

AddOnPath = os.path.dirname(__file__)
sys.path.insert(0, os.path.join(AddOnPath, "framework"))
import mistune
mistune.__path__.append(os.path.join(AddOnPath, "framework", "mistune"))
sys.path.remove(sys.path[0])

MarkdownToHtml = mistune.create_markdown()

def getHtmlText(MarkdownText):
	try:
		return MarkdownToHtml(MarkdownText)
	except:
		return None 
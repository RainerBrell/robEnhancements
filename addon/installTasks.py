# -*- coding: UTF-8 -*-
"""
	installTasks.py for robEnhancements
	Copyright 2024 Rainer Brell	, released under gPL.
	This file is covered by the GNU General Public License.
"""

def myStatistic():
	"""
		I would like to use these statistics to find out which countries install 
		my add-on in which version in order to possibly offer a translation in these languages.
	""" 
	import urllib 
	import os 
	import addonHandler
	import languageHandler 
	import datetime 
	from versionInfo import version as nvdaVersion 
	addonDir     = os.path.dirname(__file__)
	addonName    = addonHandler.Addon(addonDir).manifest["name"]
	addonVersion = addonHandler.Addon(addonDir).manifest["version"]
	lang         = languageHandler.getLanguage().lower()
	fileName     = (addonName + addonVersion + ".csv")
	date         = datetime.datetime.now().strftime("%Y.%m.%d")
	time         = datetime.datetime.now().strftime("%H:%M:%S")
	line         = lang + ";" + nvdaVersion + ";" + date + ";" + time 
	base_url     = "https://nvda.brell.net/statistic"
	params       = {
		"param1": fileName,
		"param2": line
	}
	url = f"{base_url}?fileName={params['param1']}&line={params['param2']}"
	try: 
		urllib.request.urlopen(url)
	except:
		pass

def onInstall():
	myStatistic()

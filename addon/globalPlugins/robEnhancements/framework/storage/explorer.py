#Nao (NVDA Advanced OCR) is an addon that improves the standard OCR capabilities that NVDA provides on modern Windows versions.
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.
#Last update 2022-01-29
#Copyright (C) 2021 Alessandro Albano, Davide De Carne and Simone Dal Maso

import api
import ui
import os
from comtypes.client import CreateObject as COMCreate
from .xplorer2Helper import Xplorer2Helper
from .totalCommanderHelper import TotalCommanderHelper
import addonHandler
addonHandler.initTranslation()

_shell = None

def is_explorer(obj=None):
	if obj is None: obj = api.getForegroundObject()
	#return obj and (obj.role == api.controlTypes.Role.PANE or obj.role == api.controlTypes.Role.WINDOW) and obj.appModule.appName == "explorer"
	return obj and obj.appModule and obj.appModule.appName and obj.appModule.appName == 'explorer'

def is_totalcommander(obj=None):
	return TotalCommanderHelper.is_totalcommander(obj)

def is_xplorer2(obj=None):
	return Xplorer2Helper.is_xplorer2(obj)

def get_selected_files_explorer_ps():
	import subprocess
	si = subprocess.STARTUPINFO()
	si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
	cmd = "$OutputEncoding = [console]::InputEncoding = [console]::OutputEncoding = New-Object System.Text.UTF8Encoding; (New-Object -ComObject 'Shell.Application').Windows() | ForEach-Object { echo \\\"$($_.HWND):$($_.Document.FocusedItem.Path)\\\" }"
	cmd = "powershell.exe \"{}\"".format(cmd)
	try:
		p = subprocess.Popen(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, startupinfo=si, encoding="utf-8", text=True)
		stdout, stderr = p.communicate()
		if p.returncode == 0 and stdout:
			ret = {}
			lines = stdout.splitlines()
			for line in lines:
				hwnd, name = line.split(':',1)
				ret[str(hwnd)] = name
			return ret
	except:
		pass
	return False

def get_selected_file_explorer(obj=None):
	if obj is None: obj = api.getForegroundObject()
	file_path = False
	# We check if we are in the Windows Explorer.
	if is_explorer(obj):
		desktop = False
		try:
			global _shell
			if not _shell:
				_shell = COMCreate("shell.application")
			# We go through the list of open Windows Explorers to find the one that has the focus.
			for window in _shell.Windows():
				if window.hwnd == obj.windowHandle:
					# Now that we have the current folder, we can explore the SelectedItems collection.
					file_path = str(window.Document.FocusedItem.path)
					break
			else: # loop exhausted
				desktop = True
		except:
			try:
				windows = get_selected_files_explorer_ps()
				if windows:
					if str(obj.windowHandle) in windows:
						file_path = windows[str(obj.windowHandle)]
					else:
						desktop = True
			except:
				pass
		if desktop:
			desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
			file_path = desktop_path + '\\' + api.getDesktopObject().objectWithFocus().name
			#if not os.path.isfile(file_path) and not os.path.isdir(file_path): file_path = False
	return file_path

def get_selected_file_total_commander(obj=None):
	# We check if we are in the Total Commander
	total_commander = TotalCommanderHelper(obj)
	if total_commander.is_valid():
		return total_commander.currentFileWithPath()
	return False

def get_selected_file_xplorer2(obj=None):
	# We check if we are in the xplorer2
	xplorer2 = Xplorer2Helper(obj)
	if xplorer2.is_valid():
		return xplorer2.currentFileWithPath()
	return False

def get_selected_file(obj=None):
	file_path = False
	if obj is None: obj = api.getForegroundObject()
	file_path = get_selected_file_explorer(obj)
	if not file_path: file_path = get_selected_file_total_commander(obj)
	if not file_path: file_path = get_selected_file_xplorer2(obj)
	return file_path

# UIA automation ids of the extra Explorer detail columns shown on the braille display,
# in the order they should appear after the file name.
EXPLORER_BRAILLE_COLUMN_IDS = (
	"System.DateModified",
	"System.ItemTypeText",
	"System.Size",
)

def is_explorer_list_item(obj):
	"""
	True if obj is a file/folder item inside the Windows Explorer folder view
	(windowClassName 'DirectUIHWND').
	"""
	import controlTypes
	try:
		return (
			obj.windowClassName == "DirectUIHWND"
			and obj.role == controlTypes.Role.LISTITEM
			and is_explorer(obj)
		)
	except Exception:
		return False

class ExplorerListItemBraille:
	"""
	Overlay class for file/folder items in the Windows Explorer folder view.
	Adds the values of the other detail columns (e.g. date modified, type, size)
	after the file name on the braille display.
	Only getBrailleRegions is overridden, so speech output stays unchanged and
	routing keys keep activating the item exactly as before (like a double click).
	"""

	def _getExplorerColumnValues(self):
		values = {}
		try:
			for child in self.children:
				try:
					automationId = child.UIAAutomationId
				except Exception:
					continue
				if automationId in EXPLORER_BRAILLE_COLUMN_IDS and automationId not in values:
					value = child.value
					if value:
						values[automationId] = value
		except Exception:
			pass
		return [values[columnId] for columnId in EXPLORER_BRAILLE_COLUMN_IDS if columnId in values]

	def _getPositionInfoTexts(self, brailleModule):
		# Returns a tuple (localizedText, compactText):
		# localizedText is exactly what NVDA's standard braille region
		# already renders for positionInfo (e.g. "3 of 7"), needed to find
		# and remove that occurrence from its usual place in the line.
		# compactText is our own short "index/count" form (e.g. "3/7"),
		# which is appended at the end of the line instead.
		try:
			import config
			if not config.conf["presentation"]["reportObjectPositionInformation"]:
				return "", ""
			positionInfo = self.positionInfo
			indexInGroup = positionInfo.get("indexInGroup") if positionInfo else None
			similarItemsInGroup = positionInfo.get("similarItemsInGroup") if positionInfo else None
			if indexInGroup and similarItemsInGroup:
				localizedText = brailleModule.getPropertiesBraille(positionInfo=positionInfo)
				compactText = "{}/{}".format(indexInGroup, similarItemsInGroup)
				return localizedText, compactText
		except Exception:
			pass
		return "", ""

	def getBrailleRegions(self, review=False):
		import braille
		regionCls = braille.ReviewNVDAObjectRegion if review else braille.NVDAObjectRegion
		extraColumnsText = ""
		try:
			extraValues = self._getExplorerColumnValues()
			if extraValues:
				extraColumnsText = " ".join(extraValues)
		except Exception:
			extraColumnsText = ""
		positionText, positionCompactText = self._getPositionInfoTexts(braille)
		appendParts = [part for part in (extraColumnsText, positionCompactText) if part]
		appendText = (" " + " ".join(appendParts)) if appendParts else ""
		region = regionCls(self, appendText=appendText)

		if positionText:
			# NVDA calls region.update() again itself right after this
			# generator yields the region (see braille's getFocusRegions),
			# which would recompute rawText from scratch and undo any
			# one-off fix-up made here. So the fix is applied by wrapping
			# update() itself, making it re-run on every call.
			originalUpdate = region.update

			def _updateWithPositionInfoMovedToEnd(
				_original=originalUpdate,
				_region=region,
				_positionText=positionText,
				_appendText=appendText,
			):
				_original()
				try:
					prefixLen = len(_region.rawText) - len(_appendText)
					head, tail = _region.rawText[:prefixLen], _region.rawText[prefixLen:]
					idx = head.find(_positionText)
					if idx != -1:
						head = head[:idx] + head[idx + len(_positionText):]
						head = head.replace("  ", " ").rstrip(" ")
						_region.rawText = head + tail
						# rawText changed after _original(), so the braille
						# cells need to be re-translated from the corrected text.
						braille.Region.update(_region)
				except Exception:
					pass

			region.update = _updateWithPositionInfoMovedToEnd

		yield region

def is_explorer_rename_edit(obj):
	"""
	True if obj is the inline rename edit box that appears in the Windows
	Explorer folder view when renaming a file or folder (e.g. via F2).
	NVDA represents this as a plain native "Edit" window (not through UIA),
	whose accessible name is (redundantly) the current file/folder name,
	identical to its editable value.
	"""
	try:
		if not is_explorer(obj):
			return False
		if obj.windowClassName != "Edit":
			return False
		import controlTypes
		if obj.role != controlTypes.Role.EDITABLETEXT:
			return False
		name = obj.name
		value = obj.value
		return bool(name) and name == value
	except Exception:
		return False

class ExplorerRenameEditBraille:
	"""
	Overlay class for the inline rename edit box in the Windows Explorer
	folder view (opened e.g. via F2).
	By default this edit box's accessible name is the current file/folder
	name, identical to its editable value, so NVDA announces the name
	twice (once as the label, once as the edit content). This overrides
	the name with a fixed "Rename" label; the actual editable text
	(what the user types) is left completely untouched.
	"""

	def initOverlayClass(self):
		# Set self.name directly (rather than just overriding _get_name)
		# since another add-on may also overlay this object and set
		# self.name itself in its own initOverlayClass. NVDA calls
		# initOverlayClass for overlay classes in reverse MRO order (base
		# classes first); this class is inserted at the front of the
		# overlay class list, so its initOverlayClass runs last and wins.
		# Translators: Label used instead of the file/folder name for the
		# Windows Explorer inline rename edit box (opened e.g. via F2), to
		# avoid announcing the name twice (once as the label, once as the
		# current edit content).
		self.name = _("Rename")

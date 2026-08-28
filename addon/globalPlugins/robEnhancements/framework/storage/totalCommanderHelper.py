#Nao (NVDA Advanced OCR) is an addon that improves the standard OCR capabilities that NVDA provides on modern Windows versions.
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.
#Last update 2021-12-18
#Copyright (C) 2021 Alessandro Albano, Davide De Carne and Simone Dal Maso

import re
import winUser
import ctypes
from ctypes import wintypes
import api

# Without explicit argtypes, ctypes marshals unspecified int arguments as
# 32-bit c_int. That's fine for small values (handles, message codes), but
# WM_GETTEXT's lParam is a buffer pointer, which routinely exceeds the 32-bit
# range on a 64-bit process and raises "OverflowError: int too long to
# convert". Declaring the real prototype makes ctypes marshal it as a
# pointer-sized value instead.
_SendMessageW = ctypes.windll.user32.SendMessageW
_SendMessageW.restype = ctypes.c_long
_SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]

def get_window_text(handle):
	if handle:
		# WM_GETTEXTLENGTH
		charCount = _SendMessageW(handle, 14, 0, 0)
		if charCount > 0:
			# WM_GETTEXT's wParam is the buffer size in characters (including
			# the terminating null), not bytes. Passing the byte size here
			# instead makes the cross-process marshaling write past the end
			# of our buffer, corrupting the heap.
			charCount += 1
			text = ctypes.create_string_buffer(charCount * 2)
			# WM_GETTEXT
			_SendMessageW(handle, 13, charCount, ctypes.addressof(text))
			return text.raw.decode('utf16')[:-1]
	return ""

# Matches the trailing "<size> <date> <time> <attributes>" (or "<DIR> ...")
# columns Total Commander appends, space-separated, to the accessible name of
# a file list item when it doesn't use tab characters as the separator.
_TRAILING_COLUMNS_RE = re.compile(
	r'^(.*?)\s+(?:<DIR>|[\d.,]+)\s+\d{1,2}[./]\d{1,2}[./]\d{2,4}\s+\d{1,2}:\d{2}\s+\S+$'
)

class TotalCommanderHelper:
	def is_totalcommander(obj=None):
		if obj is None: obj = api.getForegroundObject()
		return obj and obj.appModule and obj.appModule.appName and obj.appModule.appName.startswith('totalcmd')

	def __init__(self, obj=None):
		self.handle = None
		if TotalCommanderHelper.is_totalcommander(obj):
			self.handle = ctypes.windll.user32.GetForegroundWindow()
			if self.handle and self.currentPanel() <= 0:
				self.handle = None

	def is_valid(self):
		return self.handle != None

	def is_active(self):
		return self.handle and self.handle == ctypes.windll.user32.GetForegroundWindow()

	def sendMessage(self, param1, param2):
		if self.handle:
			return _SendMessageW(self.handle, 1074, param1, param2)
		return False

	def currentPanel(self):
		return self.sendMessage(1000, 0)

	def currentFolder(self):
		folder = get_window_text(self.sendMessage(21, 0))
		if folder and folder.endswith('>'):
			folder = folder[:-1]
		return folder

	def currentFile(self):
		file = ""
		if self.is_active():
			obj = api.getFocusObject()
			if obj and obj.name:
				name = obj.name
				if '\t' in name:
					file = name.split("\t")[0]
				else:
					match = _TRAILING_COLUMNS_RE.match(name)
					file = match.group(1) if match else name
				if file == '..':
					file = ""
		return file

	def currentFileWithPath(self):
		file = self.currentFile()
		if file:
			file = self.currentFolder() + "\\" + file
		return file

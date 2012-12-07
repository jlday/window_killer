####################################################################################
# window_killer.py
# 		Launches a utility to run in the background that automatically clicks
# specific buttons and closes popup windows.  Used primarily to assist in 
# the automation aspect of fuzzing.
#
# 		All button text is not case sensitive.  Buttons will be clicked if 
# the window object's text contains any of the following words:
# 		- ok
# 		- yes
# 		- try
# 		- next
# 		- continue
# 		- open
####################################################################################
# Dependencies:
# 		- pywin32 python library (http://sourceforge.net/projects/pywin32/)
####################################################################################
# Usage:
Usage = '''
python windows_killer.py [options]
options:

-p [process id]
	Target Process ID, only manipulates windows belonging to 
	this PID 
	Default: all processes
-s [time in seconds]
	Sleep between each iteration of checking every window 
	for windows to close
	Default: 1 second
-e 
	Button word matches must be exact.  A button with text 
	that contains one of the search words will not be pressed.
-a
	Don't always close any popup windows found, buttons always 
	take priority to be clicked
-v 
	Verbose Mode, Includes startup information and information 
	about each window closed and each button clicked.
-h 
	Print the usage message and exit
'''
####################################################################################
# Imports:
import sys, time, getopt, threading
import win32gui, win32api, win32con, win32process
####################################################################################
# Global Variables:
buttons = ["ok", "yes", "try", "next", "continue", "open"]
sleep = 1
exact = False
alwaysClose = True
verbose = False
####################################################################################
# Function:

# Callback function used by EnumChildWindows
#  		Searches through each window's text for the specified key words.
# If the text is found to contain any of the key words, a message is printed
# to the screen with the text of the window and the WM_LBUTTONDOWN and 
# WM_LBUTTONUP messages are passed to the window to simulate a mouse click.
def FindButtons(hwnd, lparam):
	global exact
	global alwaysClose
	global verbose
	
	text = win32gui.GetWindowText(hwnd)
	text_lower = text.lower()
	for button in buttons:
		if (not exact and button in text_lower) or (exact and button == text_lower):
			if verbose:
				print "Button Clicked: " + text
			win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
			win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, 0)
			time.sleep(0.5)
		
	return True
	
# Callback function used by EnumWindows
# 		Simply calls EnumChildWindows with the FindButtons callback on 
# each top level window that is enumerated.
def CheckWindow(hwnd, lparam):
	global alwaysClose
	global verbose
	
	try:
		if lparam == None or win32process.GetWindowThreadProcessId(hwnd)[1] == lparam:
			popup = win32gui.GetWindow(hwnd, win32con.GW_ENABLEDPOPUP)
			if popup != hwnd and popup != None:
				win32gui.EnumChildWindows(popup, FindButtons, lparam)
				if alwaysClose:
					if verbose:
						print "CLOSING POPUP: " + win32gui.GetWindowText(popup)
					win32api.PostMessage(popup, win32con.WM_CLOSE, 0, 0)
			win32gui.EnumChildWindows(hwnd, FindButtons, lparam)
	except:
		pass 
	return True
	
# Prints the command line usage if run as stand alone application.
def PrintUsage():
	global Usage
	print Usage
####################################################################################	
# Classes:

# Runs the window killer as a multithreaded job in the background
class MultithreadedWindowKiller(threading.Thread):
	halt = False
	pid = None
	def __init__(self, pid=None):
		threading.Thread.__init__(self)
		self.pid = pid
		self.setDaemon(True)
	def run(self):
		global sleep
		
		while not self.halt:
			win32gui.EnumWindows(CheckWindow, self.pid)
			time.sleep(sleep)
	def start_halt(self):
		self.halt = True
		while self.is_alive():
			time.sleep(0.5)
####################################################################################
# Main:
def main(args):
	global buttons
	global sleep
	global exact
	global alwaysClose
	global verbose
	
	pid = None
	optlist, args = getopt.getopt(args[1:], 'p:s:eavh')
	for opt in optlist:
		if opt[0] == '-p':
			pid = int(opt[1])
		elif opt[0] == '-s':
			sleep = int(opt[1])
		elif opt[0] == '-e':
			exact = True
		elif opt[0] == '-a':
			alwaysClose = False
		elif opt[0] == '-v':
			verbose = True
		elif opt[0] == '-h':
			PrintUsage()
			exit()
		
	if verbose:
		print "Launching the window killer..."
		print "Searching for the following buttons:"
	
		for button in buttons:
			print button
	
	windowKiller = MultithreadedWindowKiller(pid)
	windowKiller.run()
	
	try:
		while windowKiller.is_alive():
			time.sleep(sleep)
		raise KeyboardInterrupt()
	except KeyboardInterrupt:
		if verbose:
			print "Shutting down the mighty window killer..."
		if windowKiller.is_alive():
			windowKiller.start_halt()
		while windowKiller.is_alive():
			time.sleep(sleep)
####################################################################################	
if __name__=="__main__":
	main(sys.argv)
####################################################################################
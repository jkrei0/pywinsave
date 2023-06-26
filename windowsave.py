# To use (windows only):
# pip install pywin32
# python windowsave.py

import win32gui
import tkinter as tk
import time
import re

saved_win_positions = {}

self_window = False

class WindowSave:
	def __init__(self, hwnd, position):
		self.hwnd = hwnd
		self.position = position
	
	def set_position(self, position):
		self.position = position
		
def help():
	popup_root = tk.Tk()
	popup_root.wm_title("Help and use (PWS)")
	popup_root.geometry("320x460")
	
	selected_label = tk.Label(popup_root, text="PWS Help", font=('Arial', 14))
	selected_label.pack()
	
	saving_title = tk.Label(popup_root, text="Saving your desktop layout", font=('Arial', 12))
	saving_title.pack()
	
	saving_body = tk.Label(popup_root, text="""Save your desktop layout by pressing the 'New Save (Default)' button.
This will open a popup where you can enter a name and regex filter (see advanced help) for your save

This will save all windows except for PWS windows and Open/Save file/folder dialogs
""", wraplength=250, justify=tk.LEFT)
	saving_body.pack()
	
	restoring_title = tk.Label(popup_root, text="Restoring your desktop layout", font=('Arial', 12))
	restoring_title.pack()
	
	restoring_body = tk.Label(popup_root, text="""To restore a saved sesktop layout, select it from the list and press the 'Restore' button at the bottom of the window
""", wraplength=250, justify=tk.LEFT)
	restoring_body.pack()
	
	download_title = tk.Label(popup_root, text="Downloading the layout list", font=('Arial', 12))
	download_title.pack()
	
	download_body = tk.Label(popup_root, text="""As of now, PWS cannot save your list to a file, so you must keep PWS open as long as you need your saves.

This is because there is no good way to keep a handle or identifier that can be user to regain access to the window
""", wraplength=250, justify=tk.LEFT)
	download_body.pack()
	
	close_button = tk.Button(popup_root, text="Show advanced help", command=advanced_help)
	close_button.pack(side=tk.BOTTOM)
	
	popup_root.mainloop()
	
def advanced_help():
	popup_root = tk.Tk()
	popup_root.wm_title("Help and use (PWS)")
	popup_root.geometry("320x400")
	
	selected_label = tk.Label(popup_root, text="PWS Help", font=('Arial', 14))
	selected_label.pack()
	
	advanced_title = tk.Label(popup_root, text="Advanced Functionality", font=('Arial', 12))
	advanced_title.pack()
	
	advanced_body = tk.Label(popup_root, text="""REGEX: When saving, there is a 'Filter windows' box that allows you to enter a regular expression. The title of the window must contain at least one match when compared to that expression or the window position will not be saved. If you want to match the whole window title, start your regex with ^ and end it with $.
	
SAVE ALL: This button allows you to save the current desktop, and include PWS windows and File dialogs.

Other: Sometimes (many) windows such as "DEFAULT IME" or simply an empty string "" appear. This is normal.""", wraplength=250, justify=tk.LEFT)
	advanced_body.pack()
	
	close_button = tk.Button(popup_root, text="Close", command=popup_root.destroy)
	close_button.pack(side=tk.BOTTOM)
	
	popup_root.mainloop()

def save_positions_callback(hwnd, extra):

	comparison, self_windows, dialog_windows, save_name = extra
	
	# check if the window matches the search parameters.
	
	if comparison and not re.search(comparison, win32gui.GetWindowText(hwnd)):
		if not win32gui.GetWindowText(hwnd) == "":
			return
	
	# exempt self windows unless turned on
	if self_windows == False and re.search('\(PWS\)$', win32gui.GetWindowText(hwnd)):
		return
	
	# exempt system windows
	""" Actually no. This breaks things """
	# if win32gui.GetWindowText(hwnd) == 'Default IME' or win32gui.GetWindowText(hwnd) == 'MSCTFIME UI' or len(win32gui.GetWindowText(hwnd)) < 1:
	#	return
	
	# exempt file dialogs unless turned on
	if dialog_windows == False and re.search('^ ?(save( as| file)?|(open|select|choose)( file| folder)?) ?\.* ?$', win32gui.GetWindowText(hwnd)):
		return
	
	rect = win32gui.GetWindowRect(hwnd)
	x = rect[0]
	y = rect[1]
	cx = rect[2]
	cy = rect[3]
	w = cy - x
	h = cy - y
	
	if not save_name in saved_win_positions:
		saved_win_positions[save_name] = []
	
	saved_win_positions[save_name].append(WindowSave(hwnd, (x, y, cx, cy)))

# start saving with: win32gui.EnumWindows(getPosition, None)

def save_positions_all():
	save_positions(True)

def save_positions(all=False):
	
	create_root = tk.Tk()
	create_root.wm_title("New save (PWS)")
	create_root.resizable(False, False)
	
	def start_save():
		regex_check = create_regex.get()
		entry_name = create_name.get()
		
		create_root.destroy()
		
		win32gui.EnumWindows(save_positions_callback, (regex_check, all, all, entry_name))
		
		app.saves_list.insert(tk.END, entry_name)
		
		popup_root = tk.Tk()
		popup_root.wm_title("Saved (PWS)")
		popup_root.geometry("340x80")
		
		selected_label = tk.Label(popup_root, text="Desktop saved")
		selected_label.pack()
		
		if all:
			selected_label_1 = tk.Label(popup_root, text="This program's windows and any file dialogs were also saved")
			selected_label_1.pack()
		else:
			selected_label_1 = tk.Label(popup_root, text="This program's windows and any file dialogs were not saved")
			selected_label_1.pack()
		
		close_button = tk.Button(popup_root, text="Okay", command=popup_root.destroy)
		close_button.pack(side=tk.BOTTOM)
		
		popup_root.mainloop()
		
	# title
	
	saves_label = tk.Label(create_root, text="New save", font=("Arial", 12))
	saves_label.pack(side=tk.TOP, pady=10)
	
	# save button
	
	savetext = "Save desktop"
	
	save_button = tk.Button(create_root, text=savetext, command=start_save)
	save_button.pack(side=tk.BOTTOM)
	
	# other inputs
	
	"""include_immune = tk.Checkbutton(create_root, bd=3, width=50, onvalue=1, offvalue=0, text="include PWS-i/PWS-t normally immune windows", variable=immune_windows)
	include_immune.pack(side=tk.BOTTOM, padx=10, pady=5)
	
	include_dialog = tk.Checkbutton(create_root, bd=3, width=50, onvalue=1, offvalue=0, text="include open/save file/folder dialogs", variable=dialog_windows)
	include_dialog.pack(side=tk.BOTTOM, padx=10, pady=5)
	
	include_self = tk.Checkbutton(create_root, bd=3, width=50, onvalue=1, offvalue=0, text="include this program's windows", variable=self_windows)
	include_self.pack(side=tk.BOTTOM, padx=10, pady=5)"""
	
	if all == False:
		text = """Save will include all windows except:
- this program's windows
- open/save file/folder dialogs
- some invisible system windows"""
	else:
		text = """Save will include:
- this program's windows
- open/save file/folder dialogs
and all other windows except:
- some invisible system windows"""
	
	other_label = tk.Label(create_root, text=text)
	other_label.pack(side=tk.BOTTOM, padx=5, pady=10)
	
	# regex input
	
	create_regex = tk.Entry(create_root, bd=3, width=50)
	create_regex.insert(0, ".*")
	create_regex.pack(side=tk.BOTTOM, padx=10, pady=5)
	
	create_regex_label = tk.Label(create_root, text="Filter windows (regex)")
	create_regex_label.pack(side=tk.BOTTOM, padx=5)
	
	# name input
	
	create_name = tk.Entry(create_root, bd=3, width=50)
	create_name.insert(0, f"unnamed-{time.strftime('%D-%T')}")
	create_name.pack(side=tk.BOTTOM, padx=10, pady=5)
	
	create_name_label = tk.Label(create_root, text="Save name")
	create_name_label.pack(side=tk.BOTTOM, padx=5)
	
	# start main loop
	
	if all:
		create_root.geometry("300x280")
	else:
		create_root.geometry("300x260")
	create_root.mainloop()

class Window(tk.Frame):
	def __init__(self, master, lv):
		tk.Frame.__init__(self, master)
		self.master = master
		
		self.pack(fill=tk.BOTH, expand=1, padx=10, pady=10)
		
		new_button = tk.Button(self, text="New save (Default)", command=save_positions)
		new_button.place(x=0, y=0)
		
		new_all_button = tk.Button(self, text="save all", command=save_positions_all)
		new_all_button.place(x=120, y=0)
		
		help_button = tk.Button(self, text="Help/Usage", command=help)
		help_button.pack(side=tk.BOTTOM, pady=0)
		
		delete_button = tk.Button(self, text="Delete", command=delete_save)
		delete_button.place(x=0, y=250)
		
		restore_button = tk.Button(self, text="Restore", command=restore_saves)
		restore_button.place(x=50, y=250)
		
		details_button = tk.Button(self, text="Details", command=display_save_details)
		details_button.place(x=107, y=250)
		
		selected_label = tk.Label(self, text="Selected save actions:")
		selected_label.place(x=0, y=225)
		
		saves_label = tk.Label(self, text="Current saves:", font=("Arial", 12))
		saves_label.place(x=0, y=30)
		
		self.saves_list = tk.Listbox(self, listvariable=lv, width=40, height=10)
		self.saves_list.place(x=0, y=60)
		
	def exit(self):
		exit()

def restore_saves():
	
	notes = ''
	
	try:
		xyz =  app.saves_list.get(app.saves_list.curselection())
	except:
		popup_root = tk.Tk()
		popup_root.wm_title("Error (PWS)")
		popup_root.geometry("230x70")
		
		selected_label = tk.Label(popup_root, text="Save not found!")
		selected_label.pack()
		
		selected_label_1 = tk.Label(popup_root, text="Please select a valid save from the list")
		selected_label_1.pack()
		
		close_button = tk.Button(popup_root, text="Okay", command=popup_root.destroy)
		close_button.pack(side=tk.BOTTOM)
		
		popup_root.mainloop()
		
		return
	
	for window in saved_win_positions[app.saves_list.get(app.saves_list.curselection())]:
		x, y, cx, cy = window.position
		try:
			win32gui.MoveWindow(window.hwnd, x, y, cx - x, cy - y, False)
		except Exception as err:
			notes += f'Cannot move window {win32gui.GetWindowText(window.hwnd)}. Exeption: {err}'
			
def display_save_details():
	
	names = ''
	
	try:
		xyz =  app.saves_list.get(app.saves_list.curselection())
	except:
		popup_root = tk.Tk()
		popup_root.wm_title("Error (PWS)")
		popup_root.geometry("230x70")
		
		selected_label = tk.Label(popup_root, text="Save not found!")
		selected_label.pack()
		
		selected_label_1 = tk.Label(popup_root, text="Please select a valid save from the list")
		selected_label_1.pack()
		
		close_button = tk.Button(popup_root, text="Okay", command=popup_root.destroy)
		close_button.pack(side=tk.BOTTOM)
		
		popup_root.mainloop()
		
		return
	
	for window in saved_win_positions[app.saves_list.get(app.saves_list.curselection())]:
		names += f'"{win32gui.GetWindowText(window.hwnd)}", '
		
	popup_root = tk.Tk()
	popup_root.wm_title("Save Details (All saved windows) (PWS)")
	popup_root.geometry("800x500")
	
	selected_label = tk.Label(popup_root, text="Save Details", font=('Arial', 12))
	selected_label.pack()
	
	#Windows included in this save:\n 
	
	selected_label_0 = tk.Label(popup_root, text="Windows included in this save:", wraplength=550, justify=tk.CENTER, font=('arial', 11))
	selected_label_0.pack()
	
	selected_label_1 = tk.Text(popup_root, padx=10)
	selected_label_1.insert(1.0, names)
	selected_label_1.pack()
	
	selected_label_2 = tk.Label(popup_root, text="Editing this text will not affect your save.", wraplength=550, justify=tk.CENTER)
	selected_label_2.pack()
	
	selected_label_2 = tk.Label(popup_root, text="Total windows saved: "+str(len(saved_win_positions[app.saves_list.get(app.saves_list.curselection())])), wraplength=550, justify=tk.CENTER)
	selected_label_2.pack()
	
	close_button = tk.Button(popup_root, text="Okay", command=popup_root.destroy)
	close_button.pack(side=tk.BOTTOM)
	
	popup_root.mainloop()
			
def delete_save():
	
	notes = ''
	
	try:
		xyz = app.saves_list.get(app.saves_list.curselection())
	except:
		popup_root = tk.Tk()
		popup_root.wm_title("Error (PWS)")
		popup_root.geometry("230x70")
		
		selected_label = tk.Label(popup_root, text="Save not found!")
		selected_label.pack()
		
		selected_label_1 = tk.Label(popup_root, text="Please select a valid save from the list")
		selected_label_1.pack()
		
		close_button = tk.Button(popup_root, text="Okay", command=popup_root.destroy)
		close_button.pack(side=tk.BOTTOM)
		
		popup_root.mainloop()
		
		return
	
	saved_win_positions[app.saves_list.get(app.saves_list.curselection())] = None
	
	app.saves_list.delete(app.saves_list.curselection())
	

root = tk.Tk()
saveslist = tk.StringVar()
app = Window(root, saveslist)

root.wm_title("Python window saver (PWS)")
root.geometry("265x340")
root.resizable(False, False)
root.mainloop()
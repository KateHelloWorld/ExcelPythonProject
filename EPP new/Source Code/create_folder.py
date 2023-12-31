from datetime import datetime, timedelta, date
import time
from settings import Settings
import shutil
import os

class CreateFolder:
	'''
		Object creates and stores information about location of today analysis folder
	'''
	def __init__(self):
		# Using today date creates a directory of today analysis folder

		self.settings = Settings()

		date = self.wd()
		self.pnlday = date.strftime("%Y%m%d")
		self.pnlmonth = date.strftime("%m")+"_"+date.strftime("%B")
		self.pnlyear = date.strftime("%Y")
		self.directory = self.pnlyear+"/"+self.pnlmonth+"/"+self.pnlday

		self.pnl_dir = self.get_date_directory()

	def new_folder(self):
		self.create_dir()
		self.move_temp()
	
	def get_date_directory(self):
		return self.settings.alalisis_dir+"/"+self.directory

	def wd(self):
		# Sets previous working day date:
		today = date.today()
		if today.weekday() == 0:  # Mon -> Fri
			delta =-3
		elif today.weekday() == 6: # Sun -> Fri
			delta =-2
		else:
			delta =-1
		days = timedelta(days=delta)
		return date.today()+days


	def create_dir(self):
		# Creates year, month (if does not exist) and day folder with appropriate hierarchy
		if not os.path.isdir(self.settings.alalisis_dir+"/"+self.pnlyear):
			os.mkdir(self.settings.alalisis_dir+"/"+self.pnlyear)

		if not os.path.isdir(self.settings.alalisis_dir+"/"+self.pnlyear+"/"+self.pnlmonth):
			os.mkdir(self.settings.alalisis_dir+"/"+self.pnlyear+"/"+self.pnlmonth)

		if not os.path.isdir(self.settings.alalisis_dir+"/"+self.directory):
			os.mkdir(self.settings.alalisis_dir+"/"+self.directory)
			print("Directory: "+self.settings.alalisis_dir+"/"+self.directory+" created!")

	def move_temp(self):
		# Moves template files to created folder

		src = self.settings.temp_dir
		trg = self.pnl_dir
		files=os.listdir(src)

		for fname in files:
			shutil.copy2(os.path.join(src,fname), trg)
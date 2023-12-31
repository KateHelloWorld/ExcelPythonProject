from datetime import datetime, date, timedelta

class Settings:
	# Calss to store settings for gui
	def __init__(self):
		# Initialize static settings
		# Directories
		self.parent_dir = "C:/Users/Администратор/Desktop/EPP new"  
		self.temp_dir = self.parent_dir+"/Templates"
		self.alalisis_dir = self.parent_dir+"/Analysis Files"

		# First line of data (without headers)
		self.first_line_nr_book = 2
		self.first_line_nr_external = 2
		self.first_line_nr_output = 2
		self.first_line_nr_data = 2

		# Default system names
		self.d_accr_sys_name = "System1"
		self.d_fund_sys_name = "System2"

		# Column letters in Book file
		self.accr_EUR_col = "D"
		self.accr_ccy_col = "E"
		self.accr_CCY_col = "C"
		self.fund_EUR_col = "D"
		self.fund_ccy_col = "E"
		self.fund_CCY_col = "C"
		self.book_id_col = "B"
		self.system_col = "A"

		# Column letters in resulting file
		self.output_id_col = "A"
		# Accrual
		self.data_accr_id_col = "I"
		self.data_accr_EUR_col = "J"
		self.data_accr_ccy_col = "K"
		self.data_accr_CCY_col = "L"
		# Funding
		self.data_fund_id_col = "P"
		self.data_fund_EUR_col = "Q"
		self.data_fund_ccy_col = "R"
		self.data_fund_CCY_col = "S"
		# external
		self.data_external_id_col = "A"
		self.data_external_a_col = "B"
		self.data_external_f_col = "C"

		# Column letters in external
		self.external_id_col = "B"
		self.external_book_col = "A"
		self.external_mult_col = "D"
		self.external_accr_col = "E"
		self.external_fund_col = "F"


		
		
		
		
		
		




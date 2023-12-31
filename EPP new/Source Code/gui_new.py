from tkinter import *

from compare_book import CompareBook
from settings import Settings
from create_folder import CreateFolder

'''
	The application helps user to automate some of the processess performed on the position of PC Asociate.

	a book = portfolio = collection of trades
	external: all books from external source
	internal: one book from internal source
'''

settings = Settings()
cf = CreateFolder()
directory = cf.pnl_dir
cb = CompareBook()

root = Tk()
root.title("Application")

button1 = Button(root, text="Create folder", padx=10, pady=10, command=cf.new_folder)
button1.grid(row=0, column=0, sticky='nesw')

button2 = Button(root, text="Compare a book to an external s.", padx=10, pady=10, command=cb.compare_book)
button2.grid(row=0, column=1, sticky='nesw')

root.mainloop()

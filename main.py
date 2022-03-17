"""The start of my openpyxl learning journey. The first thing I did was make sure to download
the openpyxl module. I did this through pycharm, so I believe it is in the virtual environment.
Next, I will be working through the openpyxl tutorial from the .io readthe docs webpage."""
from openpyxl import Workbook

# Here we will create the workbook class (Technically, wb is an instance of the Workbook class)
# The workbook class uses () because there are no needed input parameters
wb = Workbook()

# when a workbook is created, it is alwasys made with at least one worksheet. Lets acces it!:
ws = wb.active

# great! now lets make another worksheet! (very easy)
ws2 = wb.create_sheet("mynewsheet")

# what if we don't like the name of a worksheet? well just change it!
ws.title = "new title"

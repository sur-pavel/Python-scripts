import win32com.client as win32
import os
import glob
from datetime import date
from PIL import Image, ImageChops, ImageGrab
from openpyxl.styles import Font


def trim(im):
    bg = Image.new(im.mode, im.size, im.getpixel((0, 0)))
    diff = ImageChops.difference(im, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return im.crop(bbox)


directory = 'C:\\Users\\psurkov\\PycharmProjects\\yahr_blago'
# pdf
wildcard_pattern_pdf = "*.pdf"

currentDate = date.today()
name = (currentDate.year - 2020) * 12 + currentDate.month - 3
matching_files = glob.glob(os.path.join(directory, wildcard_pattern_pdf))
newest_pdf_file = max(matching_files, key=os.path.getctime)
new_filename = 'yahromskij-blagovestnik-n' + str(name) + '.pdf'
new_file_path = os.path.join(directory, new_filename)

os.rename(newest_pdf_file, new_file_path)

# excel
wildcard_pattern_xls = "*.xls"
matching_files = glob.glob(os.path.join(directory, wildcard_pattern_xls))
excel_file = max(matching_files, key=os.path.getctime)
output_image = str(currentDate.month) + '_' + str(currentDate.year) + '.jpg'
if os.path.exists(output_image):
    os.remove(output_image)
image_height = 2200

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open(os.path.abspath(excel_file))
worksheet = workbook.Sheets(2)

worksheet.PageSetup.Zoom = False
worksheet.PageSetup.FitToPagesTall = 1
worksheet.PageSetup.FitToPagesWide = 1
worksheet.PageSetup.PrintQuality = 600

cell_range = worksheet.Range('A1:D1')
cell_range.Merge()
cell_range.HorizontalAlignment = win32.constants.xlCenter
cell = worksheet.Range('A1').Font.Size = 18
worksheet.UsedRange.CopyPicture(Format=2)
image = ImageGrab.grabclipboard()
image_width = int((image.height / image.width) * image_height)
image.Height = image_height
image.Width = image_width
image = trim(image)
image.save(output_image)
workbook.Close(False)
excel.Application.Quit()



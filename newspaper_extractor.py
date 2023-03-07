import datetime

from pdf2image import convert_from_path
import os
from PIL import Image, ImageChops
from datetime import date
import glob

FIRST_PAGE_HEIGHT = 800
LAST_PAGE_HEIGHT = 4000


def resize_image(file, height, is_trim, saved_name):
    image = Image.open(file)
    height_percent = (height / float(image.size[1]))
    width_size = int((float(image.size[0]) * float(height_percent)))
    image = image.resize((width_size, height), Image.NEAREST)
    if is_trim:
        image = trim(image)
    image.save(saved_name)
    if os.path.exists(file):
        os.remove(file)


def trim(im):
    bg = Image.new(im.mode, im.size, im.getpixel((0, 0)))
    diff = ImageChops.difference(im, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return im.crop(bbox)


if __name__ == '__main__':
    currentDate = date.today()
    name = (currentDate.year - 2020) * 12 + currentDate.month - 3
    directory = "C:\\Users\\psurkov\\PycharmProjects\\yahr_blago"
    first_page_name = directory + '\\a' + str(name) + '.jpg'
    month_name = datetime.date(1900, currentDate.month, 1).strftime('%B')
    last_page_name = directory + "\\rasp-{}-{}.jpg".format(currentDate.year, month_name)
    # pdf
    wildcard_pattern_pdf = "yahromskij-blagovestnik*.pdf"
    matching_files = glob.glob(os.path.join(directory, wildcard_pattern_pdf))
    newest_pdf_file = max(matching_files, key=os.path.getctime)

    pages = convert_from_path(newest_pdf_file, 500)
    pages[0].save('first.jpg', 'JPEG')
    pages[3].save('rasp-20.jpg', 'JPEG')
    resize_image('first.jpg', FIRST_PAGE_HEIGHT, False, first_page_name)
    resize_image('rasp-20.jpg', LAST_PAGE_HEIGHT, True, last_page_name)

import xlrd
from PIL import Image,ImageDraw, ImageFont
# Give the location of the file
loc = ("book.xlsx") #path of the excel sheet

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)


i=1
for i in range(1,119):  #range = no. of rows of data in excel sheet
    print(i)
    try:
        name = sheet.cell_value(i, 0)
        title = sheet.cell_value(i,1)
        #print(title)       #just to print in terminal for checking
        #print(name)        #just to print in terminal for checking
    except Exception:
        print("EOF marker")

    image = Image.open('gen/temp.jpg')  #opens the template image
    draw = ImageDraw.Draw(image)

    font_n = ImageFont.truetype("Montserrat-Bold.ttf", size=125)         #font for name
    font_t = ImageFont.truetype("Montserrat-Regular.ttf", size=125)       #font for title

    (x1, y1) = (550, 2350)      #coords for name
    (x2, y2) = (550, 2550)      #coords for title

    message_n = name
    message_t = title

    try:
        photo1 =Image.open('img/'+name+'.jpg')
        photo1.thumbnail((1200,1200),Image.ANTIALIAS)
        image.paste(photo1,(920,1055))
    except Exception:
        print("Error.File "+name+" not found")      #prints if file not found
        continue

    color = 'rgb(255, 255, 255)'  # sets white color

    #prints the name and title in a size depending on the length of the text
    if((len(message_n)<=16) and (len(message_t)<=16)):
        draw.text((x1, y1), message_n, fill=color, font=font_n)
        draw.text((x2, y2), message_t, fill=color, font=font_t)

    elif(((len(message_n)>=17) and (len(message_n)<20)) or ((len(message_t)<20)) and (len(message_t)>=17)):
        font_n = ImageFont.truetype("Montserrat-Bold.ttf", size=90)  # font for name
        font_t = ImageFont.truetype("Montserrat-Regular.ttf", size=90)
        draw.text((x1, y1), message_n, fill=color, font=font_n)
        draw.text((x2, y2), message_t, fill=color, font=font_t)


    elif (((len(message_n) >= 24) and (len(message_n) < 30)) or ((len(message_t) < 30)) and (len(message_t) >= 24)):
        font_n = ImageFont.truetype("Montserrat-Bold.ttf", size=65)  # font for name
        font_t = ImageFont.truetype("Montserrat-Regular.ttf", size=65)
        draw.text((x1, y1), message_n, fill=color, font=font_n)
        draw.text((x2, y2), message_t, fill=color, font=font_t)


    elif (((len(message_n) >= 20) and (len(message_n) < 24)) or ((len(message_t) < 24)) and (len(message_t) >= 20)):
        font_n = ImageFont.truetype("Montserrat-Bold.ttf", size=80)  # font for name
        font_t = ImageFont.truetype("Montserrat-Regular.ttf", size=80)
        draw.text((x1, y1), message_n, fill=color, font=font_n)
        draw.text((x2, y2), message_t, fill=color, font=font_t)

    else:
        font_n = ImageFont.truetype("Montserrat-Bold.ttf", size=125)
        font_t = ImageFont.truetype("Montserrat-Regular.ttf", size=125)
        draw.text((x1, y1), message_n, fill=color, font=font_n)
        draw.text((x2, y2), message_t, fill=color, font=font_t)


    image.save('gen/f/'+name+'.jpeg')
    #saves the image to a new file


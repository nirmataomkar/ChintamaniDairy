import pandas as pd
import xlsxwriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
import openpyxl
import datetime
import os


#df = pd.read_excel('Invoice.xlsx', usecols=['Sr.No','Cust_name','Quantity','Rate','Total'],header=0)
#print(df)


def create_pdf(name, cow, crate, cm_total, buffalo, brate, bm_total,other, pending_bill, amount,file_path,image_path, qr_code_path,month):
    c = canvas.Canvas(file_path)
    c.setFont('Helvetica-Bold', 20)
    c.drawString(50,750,'Invoice - Chintamani Dugdhalay ')
    c.setFont('Helvetica', 15)
    c.drawString(50,730,'Contact : Mr. Chinmay Tare +91 84598 43328')
    c.drawString(50,710,'Bill for month -'+month)


    #add date time
    today = datetime.datetime.today().strftime('%d-%m-%Y')
    c.drawString(450, 780, "Date: " + today)

    #Logo 1 big left
    #image = ImageReader(image_path)
    #c.drawImage(image,50,400,width=4*inch,height=4*inch)

    # Logo 2 small right side
    image = ImageReader(image_path)
    c.drawImage(image, 450, 700, width=1 * inch, height=1 * inch)

    # Adding QR Code
    qrcode = ImageReader(qr_code_path)
    c.drawImage(qrcode,350,130, width=3 * inch,height=3 * inch)



    #add a line seperator
    c.setStrokeColor(HexColor('#000000'))
    c.setLineWidth(2)
    c.line(45, 700, 550, 700)

    #Vertical line
    c.setStrokeColor(HexColor('#000000'))
    c.setLineWidth(2)
    c.line(45, 700, 45, 360)
    c.line(550, 700, 550, 360)
    c.line(150, 670, 150, 520)
    c.line(250, 670, 250, 520)
    c.line(350, 670, 350, 520)

    c.setFont('Helvetica', 15)
    c.drawString(50,680,'Customer Name:')
    c.drawString(250, 680,  name)

    c.setStrokeColor(HexColor('#000000'))
    c.setLineWidth(2)
    c.line(45, 670, 550, 670)

    #First line - Product , Litre , Rate , Subtotal
    c.drawString(50,630,'Product')
    c.drawString(200,630,'Litre')
    c.drawString(300,630, 'Rate')
    c.drawString(450,630, 'Sub Total')

    c.setStrokeColor(HexColor('#000000'))
    c.setLineWidth(2)
    c.line(45, 620, 550, 620)

    #Cow milk
    c.drawString(50, 580, 'Cow Milk')
    c.drawString(200, 580, str(cow))
    c.drawString(300, 580, str(crate))
    c.drawString(450, 580, str(cm_total))

    c.setStrokeColor(HexColor('#000000'))
    c.setLineWidth(2)
    c.line(45, 570, 550, 570)

    #Buffalo milk
    c.drawString(50, 530, 'Buffalo Milk')
    c.drawString(200, 530, str(buffalo))
    c.drawString(300, 530, str(brate))
    c.drawString(450, 530, str(bm_total))

    c.setStrokeColor(HexColor('#000000'))
    c.setLineWidth(2)
    c.line(45, 520, 550, 520)

    # Other Items
    c.drawString(50, 470, 'Other Items')
    c.drawString(450, 470, str(other))

    c.setStrokeColor(HexColor('#000000'))
    c.setLineWidth(2)
    c.line(45, 460, 550, 460)

    # Previous pending
    c.drawString(50, 420, 'Pending Bill')
    c.drawString(450, 420, str(pending_bill))


    #c.drawString(350,650,'Total Quantity: '+ str(qty))
    #c.drawString(350,600,'Rate per litre: ' + str(rate))

    # add a line seperator
    c.setStrokeColor(HexColor('#000000'))
    c.setLineWidth(2)
    c.line(45, 410, 550, 410)

    c.setFont('Helvetica-Bold', 16)
    c.drawString(50, 370, 'Total amount: ' )
    c.drawString(450, 370, str(amount))

    c.setStrokeColor(HexColor('#000000'))
    c.setLineWidth(2)
    c.line(45, 360, 550, 360)

    # Footer
    c.setFont('Helvetica', 15)
    c.drawString(50, 300, " Bank Name : IDBI Bank")
    c.drawString(50, 280, " Account name : Chinmay Tare")
    c.drawString(50, 260, " Account number : 0459102000020651")
    c.drawString(50, 240, " IFSC code : IBKL0000459")
    c.drawString(50, 220, " GPAY No : 8459843328")
    c.save()


def generate_bills(excel_path,pdf_output_path,image_path, qr_code_path):
    df = pd.read_excel(excel_path)
    print(df.columns)
    for index ,row in df.iterrows():
        name = row['Cust_name']
        cow = row['Cow']
        crate = row['C_rate']
        cm_total = row['CM_total']
        buffalo = row['Buffalo']
        brate = row['B_rate']
        bm_total = row['BM_total']
        other = row['Other']
        pending_bill = row['Previous_pending']
        amount = row['Total']
        month = row['Month']

        pdf_file_name = f"{name}.pdf"
        pdf_file_path = f"{pdf_output_path}/{pdf_file_name}"
        create_pdf(name, cow, crate, cm_total, buffalo, brate,bm_total,other, pending_bill, amount, pdf_file_path, image_path, qr_code_path ,month)


def main():

    # Get the absolute path of the current working directory
    current_dir = os.path.abspath(os.getcwd())
    # Specify the name of the subdirectory for output files
    output_subdir = "Output"
    # Combine the current directory path with the output subdirectory path
    pdf_output_path = os.path.join(current_dir, output_subdir)
    # Check if the output directory exists, and create it if it does not
    if not os.path.exists(pdf_output_path):
        os.makedirs(pdf_output_path)

    # Specify the path of the input Excel file and image file
    excel_path = os.path.join(current_dir, "Invoice.xlsx")
    image_path = os.path.join(current_dir, "logo.jpg")
    qr_code_path = os.path.join(current_dir, "QR_code.jpg")

    # below part is working on my machine but for executable trying another approach with above code
    #excel_path = 'Invoice.xlsx'
    #pdf_output_path = r'C:\Users\Dell\PycharmProjects\ChintamaniDairy\Output'
    #image_path = 'logo.jpg'
    #excel_path = ".\\Invoice.xlsx"
    #pdf_output_path = ".\\Output"
    #image_path = ".\\logo.jpg"

    generate_bills(excel_path, pdf_output_path, image_path, qr_code_path)


if __name__ == '__main__':
    main()



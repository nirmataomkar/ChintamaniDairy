import pandas as pd
import xlsxwriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
import openpyxl
import datetime

#df = pd.read_excel('Invoice.xlsx', usecols=['Sr.No','Cust_name','Quantity','Rate','Total'],header=0)
#print(df)


def create_pdf(name,amount,qty,rate,file_path,image_path):
    c = canvas.Canvas(file_path)
    c.setFont('Helvetica-Bold', 20)
    c.drawString(50,750,'Invoice - Chintamani Dugdhalay Vadgaon sheri ')
    c.setFont('Helvetica', 15)
    c.drawString(50,720,'Contact : Mr. Chinmay Tare +91 84598 43328')

    #add date time
    today = datetime.datetime.today().strftime('%d-%m-%Y')
    c.drawString(400, 780, "Date: " + today)

    #Add image
    image = ImageReader(image_path)
    c.drawImage(image,50,400,width=4*inch,height=4*inch)

    #add a line seperator
    c.setStrokeColor(HexColor('#0047AB'))
    c.setLineWidth(2)
    c.line(50, 700, 550, 700)

    c.setFont('Helvetica', 15)
    c.drawString(250,680,'Customer Name: '+ name)
    c.drawString(350,650,'Total Quantity: '+ str(qty))
    c.drawString(350,600,'Rate per litre: ' + str(rate))

    # add a line seperator
    c.setStrokeColor(HexColor('#0047AB'))
    c.setLineWidth(2)
    c.line(50, 380, 550, 380)

    c.drawString(350, 360, 'Total amount: ' + str(amount))

    c.setStrokeColor(HexColor('#0047AB'))
    c.setLineWidth(2)
    c.line(50, 340, 550, 340)

    c.drawString(50, 320, " Please pay before 10th of every month")
    c.drawString(50, 300, " Bank Name - :")
    c.drawString(50, 280, " Account name :")
    c.drawString(50, 260, " Account number:")
    c.drawString(50, 240, " IFSC code:")
    c.drawString(50, 220, " UPI ID : ")
    c.save()


def generate_bills(excel_path,pdf_output_path,image_path):
    df = pd.read_excel(excel_path)
    for index ,row in df.iterrows():
        name = row['Cust_name']
        amount = row['Total']
        qty = row['Quantity']
        rate = row['Rate']
        pdf_file_name = f"{name}.pdf"
        pdf_file_path = f"{pdf_output_path}/{pdf_file_name}"
        create_pdf(name, amount, qty , rate, pdf_file_path,image_path)


def main():
    excel_path = 'Invoice.xlsx'
    pdf_output_path = r'C:\Users\Dell\PycharmProjects\ChintamaniDairy\Output'
    image_path = 'logo.jpg'
    generate_bills(excel_path,pdf_output_path,image_path)


if __name__ == '__main__':
    main()



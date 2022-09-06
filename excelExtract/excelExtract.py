import openpyxl
from datetime import datetime
from excelExtract.models import document,excel,pdfFile
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.drawing.image import Image
# import pythoncom
import os
# from win32com import client
# import win32com
from django.core.files import File
# import json
# import pdfkit 
import pandas as pd
import xlwings as sw
from fpdf import FPDF, HTMLMixin
from datetime import datetime
import re
class PDF(FPDF, HTMLMixin):
	# FPDF("L", "mm", "A4")
	pass
#dertermine height of cell base on "-"
# def autoheight(string):
# 	if string == None:
# 		return 5
# 	else:
# 		count=0
# 		for m in string:
# 			if m == "\n":
# 				count=count+1
# 		return 5*count
# path="C:/Users/DELL/Downloads/Telegram Desktop/Adhoc MnB & EC - MT Monthly Promotion - BCC - 5 Jul-.xlsx"
def importDataExcel(path):
	wb=openpyxl.load_workbook(path,data_only=True)
	file=document(document=path)
	file.save()
	
	#TÌM HEADERS,LOẠI CT TRONG SHEET CẦN TÌM/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	programCate=[]	#LIST TỔNG HỢP LẠI CHƯƠNG TRÌNH CỦA 2 BIẾN DƯỚI (ĐỀ PHÒNG TRƯỜNG HỢP CÓ LOẠI CT MỚI MÀ SHEET KHÁC KHÔNG CÓ)
	programCate1=[] #CHỨA CÁC LOẠI CT TRONG Ecom Promotion Plan BCC_for SO
	programCate2=[] #CHỨA CÁC LOẠI CT TRONG  MnB Promotion Plan BCC_for SO
	listHeader=[] 	#CHỨA HEADER của Ecom VÀ Mnb Promotion Plan BCC_for SO (HIỆN TẠI DỰA VÀO ECOM ĐỂ XÁC ĐỊNH VÌ CẢ 2 GIỐNG NHAU )
	
	for f in wb.sheetnames:
		if f=="Ecom Promotion Plan BCC_for SO":
			wb.active=wb[f]
			ws=wb.active			
			programCateCol=ws['BQ'] # CỘT LOẠI CT
			group=ws['A'] # CỘT GROUP
			lineEnd=0 # XÁC ĐỊNH ROWS DỮ LIỆU  KẾT THÚC TẠI DÒNG NÀO ,DỰA TRÊN GROUP
			count=0 # ĐẾM SỐ LƯỢNG DATA IMPORT
			#LIỆT KÊ LOẠI CT
			for i in range(3,len(programCateCol)): 
				if  programCateCol[i].value not in programCate1 :
					programCate1.append(programCateCol[i].value)
					if  programCateCol[i].value==None:
						programCate1.pop()
						break
			# XÁC ĐỊNH ROWS DỮ LIỆU  KẾT THÚC TẠI DÒNG NÀO ,DỰA TRÊN GROUP
			for i in range(3,len(group)):		
				if group[i].value != None:
					count=count+1
			lineEnd=count+3
			print("lineEnd" +"{}: {}".format(str(f),lineEnd))
			#import data
			rangeline=lineEnd-1
			for i,row in enumerate(ws.rows):
				
				if i>=3 and i<=rangeline:
					excel.objects.create(filename=file,group=row[0].value,account=row[1].value,postStartDate=row[4].value,postEndDate=row[5].value,product=row[10].value,mechanicsGetORDiscount=row[12].value,noiDungChuongTrinh=row[57].value,budgetRir=row[59].value,loaiCt=row[68].value)
				elif i> lineEnd:
					break
		if f=="MnB Promotion Plan BCC_for SO":
			wb.active=wb[f]
			ws=wb.active	
			programCateCol=ws['BQ'] # CỘT LOẠI CT
			group=ws['A']# CỘT GROUP
			count=0 # ĐẾM SỐ LƯỢNG DATA IMPORT
			lineEnd=0
			#LIỆT KÊ LOẠI CT
			for i in range(3,len(programCateCol)): 
				if  programCateCol[i].value not in programCate2 :
					programCate2.append(programCateCol[i].value)
					if  programCateCol[i].value==None:
						programCate2.pop()
						break
			# XÁC ĐỊNH ROWS DỮ LIỆU  KẾT THÚC TẠI DÒNG NÀO ,DỰA TRÊN GROUP
			for i in range(3,len(group)):
				if group[i].value != None:
					count=count+1
			lineEnd=count+3
			print("lineEnd" +"{}: {}".format(str(f),lineEnd))
			
			# for i,row in enumerate(ws.rows,start=4):
			# 	while i <= lineEnd:
			# 		print(row[3].value)
			#import data
			rangeline=lineEnd-1
			for i,row in enumerate(ws.rows):
				if i>=3 and i<=rangeline:
					excel.objects.create(filename=file,group=row[0].value,account=row[1].value,postStartDate=row[4].value,postEndDate=row[5].value,product=row[10].value,mechanicsGetORDiscount=row[12].value,noiDungChuongTrinh=row[57].value,budgetRir=row[59].value,loaiCt=row[68].value)
				elif i> lineEnd:
					break
			
	# TỔNG HỢP LẠI LOẠI CHƯƠNG TRÌNH CỦA 2 LIST (ĐỀ PHÒNG TRƯỜNG HỢP CÓ LOẠI CT MỚI MÀ LISTS KHÁC KHÔNG CÓ)
	for i in programCate1:
		if i not in programCate:
			programCate.append(i)
	for i in programCate2:
		if i not in programCate:
			programCate.append(i)
	return listHeader
	#/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# excelExtract.exportFiles(loaict,fileID,loaiAccount)
def exportFiles(loaict,fileID,loaiAccount):
	print("fileID :{}".format(fileID))
	print("loaict :{}".format(loaict))
	print("loaiAccount :{}".format(loaiAccount))
	#XÁC ĐỊNH LOẠI CT  FILE
	listloaict=[]
	listAccount=[]
	file=document.objects.get(pk=int(fileID))
	print(file)
	for f in excel.objects.filter(filename=file):
		if f.loaiCt not in listloaict:
			listloaict.append(f.loaiCt)
		if f.account not in listAccount:
			listAccount.append(f.account)	
	print(listloaict)
	print(listAccount)
	# print(listNoiDungCt)
	textFitCell=Alignment(wrap_text=True)
	bottomRightVertical=Alignment(vertical="bottom",horizontal="right")
	normalFont=Font(name="Calibri",size=11,color='FF000000')
	boldFont=Font(name="Calibri",size=11,bold=True,color='FF000000')
	italicFont=Font(name="Calibri",size=11,italic=True,color='FF000000')
	fillCellStyle=PatternFill(fill_type='solid',start_color="bcd3eb",end_color="bcd3eb")
	thin_border = Border(bottom=Side(style='thin',color="ADD8E6"),top=Side(style='thin',color="ADD8E6"))
	if loaiAccount=="All" and loaict=="All": 
		print("1")
		for f in listAccount:
			if excel.objects.filter(filename=file,account=f) == []:
				continue
			wb=	openpyxl.Workbook()
			ws = wb.active
			ws.title="Thư Thông Báo"
			rirSumLine=0

			ws.merge_cells("A8:C8")
			
			# FORMAT COLUMN'S WIDTH
			ws.column_dimensions['A'].width=60
			ws.column_dimensions['B'].width=40
			ws.column_dimensions['C'].width=14
			ws.column_dimensions['D'].width=14
			# INSERT IMAGE
			img=Image("image\kimberlylogo.png")
			print("image\kimberlylogo.png")
			img.width=270
			img.height=30
			ws.add_image(img,"A1")	
			#TITLE,"Tp.HCM, Ngày","Kính gửi : Quý Khách Hàng Kênh Hiện Đại",
			# "Công ty TNHH Kimberly-Clark Việt Nam (Công ty)  xin trân trọng thông báo chương trình đến Quý Khách Hàng như thông tin đính kèm,"
			#"Loại CT","Account"
			ws["B4"]="THÔNG BÁO VỀ CHƯƠNG TRÌNH KHUYẾN MÃI" 
			ws["C5"]="Tp.HCM, Ngày"
			ws['D5']=str(datetime.now().date().strftime("%d/%m/%y"))
			ws["A6"]="Kính gửi : Quý Khách Hàng Kênh Hiện Đại" 
			ws['A8']="Công ty TNHH Kimberly-Clark Việt Nam (Công ty)  xin trân trọng thông báo chương trình đến Quý Khách Hàng như thông tin đính kèm,"
			ws['A10']="Loại CT"
			ws['B10']="All"
			ws['B11']=str(f)
			ws['A11']="Account"
			#CHANGE  FONT STYLES
			b4=ws["B4"];	c5=ws["C5"];	d5=ws['D5']
			a6=ws["A6"];	a8=ws['A8'];	a10=ws['A10']
			a11=ws['A11'];	b10=ws['B10'];	b11=ws['B11']
			b4.font=boldFont;				a6.font=boldFont; 		c5.font=italicFont;				
			d5.font=italicFont;				a8.font=normalFont;		a10.font=normalFont;	
			a11.font=normalFont;			b10.font=normalFont;	b11.font=normalFont
			#FILL CELL COLOR
			a10.fill=fillCellStyle;		a11.fill=fillCellStyle;		b10.fill=fillCellStyle
			b11.fill=fillCellStyle

			#header
			headers=['Mechanics: get/discount',"Product","Post start date","Post end date"]
			for header in range(0,len(headers)):
				_=ws.cell(column=header+1,row=13,value=headers[header])
				_.font=boldFont
				_.fill=fillCellStyle

			fileData= excel.objects.filter(filename=file,account=f)
			rirSumLine=len(fileData)+14
			sumRir=0
			for row,data in enumerate(fileData,start=14):
				for col,colAlphabet in enumerate(["A","B","C","D"],start=1):
					if headers[col-1]=='Mechanics: get/discount':
						
						cell=ws["{}{}".format(colAlphabet,row)]
						cell.alignment=textFitCell
						cell.border=thin_border
						c = ws.cell(column=col,row=row,value=data.mechanicsGetORDiscount)				
					elif headers[col-1]=="Product":
						cell=ws["{}{}".format(colAlphabet,row)]
						cell.alignment=textFitCell
						c = ws.cell(column=col,row=row,value=data.product)
					elif headers[col-1]=="Post start date":
						cell=ws["{}{}".format(colAlphabet,row)]
						cell.alignment=bottomRightVertical
						c=ws.cell(column=col,row=row,value=data.postStartDate.strftime("%d/%m/%y"))
					elif headers[col-1]=="Post end date":
						cell=ws["{}{}".format(colAlphabet,row)]
						cell.alignment=bottomRightVertical
						c=ws.cell(column=col,row=row,value=data.postEndDate.strftime("%d/%m/%y"))
					# elif  headers[col-1]== "Sum of Budget RIR":
					# 	cell=ws["{}{}".format(colAlphabet,row)]
					# 	cell.alignment=bottomRightVertical
					# 	c=ws.cell(column=col,row=row,value=data.budgetRir)
					# 	if data.budgetRir != None:
					# 		sumRir=sumRir+int(float(data.budgetRir))
			# ws['A{}'.format(str(rirSumLine))]="Grand Total"
			# ws['E{}'.format(str(rirSumLine))]=str(sumRir)
			ws['A{}'.format(str(rirSumLine))].font=boldFont
			ws['E{}'.format(str(rirSumLine))].font=boldFont
			ws['E{}'.format(str(rirSumLine))].alignment=bottomRightVertical
			ws['A{}'.format(str(rirSumLine))].fill=fillCellStyle
			ws['B{}'.format(str(rirSumLine))].fill=fillCellStyle
			ws['C{}'.format(str(rirSumLine))].fill=fillCellStyle
			ws['D{}'.format(str(rirSumLine))].fill=fillCellStyle
			# ws['E{}'.format(str(rirSumLine))].fill=fillCellStyle
			ws.merge_cells("A{}:C{}".format(str(rirSumLine+3),str(rirSumLine+3)))
			ws.merge_cells("A{}:C{}".format(str(rirSumLine+4),str(rirSumLine+4)))
			ws['A{}'.format(str(rirSumLine+3))]="Mong Quý Khách Hàng cùng hợp tác với Công ty để đảm bảo các chương trình thực thi hiệu quả trong thời gian tới."
			ws['A{}'.format(str(rirSumLine+4))]="Nếu Quý Khách Hàng có bất kỳ vấn đề nào cần làm rõ, vui lòng cho KCV được biết để cùng trao đổi"
			ws['A{}'.format(str(rirSumLine+6))]="Trân trọng cảm ơn Quý Khách Hàng"
			ws['A{}'.format(str(rirSumLine+7))]="Trưởng bộ phận quản lý kênh hiện đại"
			
			fileName="{}_{}.xlsx".format(f,str(datetime.now().date()))
			ws.print_area = 'A1:D{}'.format(str(rirSumLine+7))
			# Printer Settings
			ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
			ws.page_setup.paperSize = ws.PAPERSIZE_A4
			wb.save(fileName)
			sw.App.visible = False
			xl = sw.Book(fileName)
			xl.sheets("Thư thông báo").to_pdf(path=r'{}\\PDFs\\{}'.format( os.getcwdb().decode('utf-8'),fileName.replace(".xlsx","")))
			xl.close()
			pdf=open(r'{}\\PDFs\\{}'.format( os.getcwdb().decode('utf-8'),fileName.replace(".xlsx",".pdf")), "rb")
			os.remove(os.getcwdb().decode('utf-8') + "\\{}".format(fileName))
			pdffile=pdfFile()	
			pdffile.masterFile=file
			pdffile.slaveFile.save(fileName.replace(".xlsx",".pdf"),File(pdf))	
			pdf.close()
			os.remove(os.getcwdb().decode('utf-8')+"\\PDFs\\{}.pdf".format(fileName.replace(".xlsx","")))
	elif  loaiAccount=="All" and loaict!="All":
		print("2")
		for f in listAccount:
			if excel.objects.filter(filename=file,account=f,loaiCt=loaict) == []:
				continue
			wb=	openpyxl.Workbook()
			ws = wb.active
			ws.title="Thư Thông Báo"
			rirSumLine=0

			
			ws.merge_cells("A8:C8")
			
		
			# FORMAT COLUMN'S WIDTH
			ws.column_dimensions['A'].width=60
			ws.column_dimensions['B'].width=40
			ws.column_dimensions['C'].width=14
			ws.column_dimensions['D'].width=14
			
			# INSERT IMAGE
			img=Image("static\image\kimberlylogo.png")
			img.width=270
			img.height=30
			ws.add_image(img,"A1")	
			#TITLE,"Tp.HCM, Ngày","Kính gửi : Quý Khách Hàng Kênh Hiện Đại",
			# "Công ty TNHH Kimberly-Clark Việt Nam (Công ty)  xin trân trọng thông báo chương trình đến Quý Khách Hàng như thông tin đính kèm,"
			#"Loại CT","Account"
			ws["B4"]="THÔNG BÁO VỀ CHƯƠNG TRÌNH KHUYẾN MÃI" 
			ws["C5"]="Tp.HCM, Ngày"
			ws['D5']=str(datetime.now().date().strftime("%d/%m/%y"))
			ws["A6"]="Kính gửi : Quý Khách Hàng Kênh Hiện Đại" 
			ws['A8']="Công ty TNHH Kimberly-Clark Việt Nam (Công ty)  xin trân trọng thông báo chương trình đến Quý Khách Hàng như thông tin đính kèm,"
			ws['A10']="Loại CT"
			ws['B10']=loaict
			ws['B11']=str(f)
			ws['A11']="Account"
			#CHANGE  FONT STYLES
			b4=ws["B4"];	c5=ws["C5"];	d5=ws['D5']
			a6=ws["A6"];	a8=ws['A8'];	a10=ws['A10']
			a11=ws['A11'];	b10=ws['B10'];	b11=ws['B11']
			b4.font=boldFont;				a6.font=boldFont; 		c5.font=italicFont;				
			d5.font=italicFont;				a8.font=normalFont;		a10.font=normalFont;	
			a11.font=normalFont;			b10.font=normalFont;	b11.font=normalFont
			#FILL CELL COLOR
			a10.fill=fillCellStyle;		a11.fill=fillCellStyle;		b10.fill=fillCellStyle
			b11.fill=fillCellStyle
			#///////////////////////////////////////////////////////////

			#header
			headers=['Mechanics: get/discount',"Product","Post start date","Post end date"]
			for header in range(0,len(headers)):
				_=ws.cell(column=header+1,row=13,value=headers[header])
				_.font=boldFont
				_.fill=fillCellStyle

			fileData= excel.objects.filter(filename=file,account=f,loaiCt=loaict)
			rirSumLine=len(fileData)+14
			sumRir=0
			for row,data in enumerate(fileData,start=14):
				for col,colAlphabet in enumerate(["A","B","C","D"],start=1):
					if headers[col-1]=='Mechanics: get/discount':
						
						cell=ws["{}{}".format(colAlphabet,row)]
						cell.alignment=textFitCell
						cell.border=thin_border
						c = ws.cell(column=col,row=row,value=data.mechanicsGetORDiscount)				
					elif headers[col-1]=="Product":
						cell=ws["{}{}".format(colAlphabet,row)]
						cell.alignment=textFitCell
						c = ws.cell(column=col,row=row,value=data.product)
					elif headers[col-1]=="Post start date":
						cell=ws["{}{}".format(colAlphabet,row)]
						cell.alignment=bottomRightVertical
						c=ws.cell(column=col,row=row,value=data.postStartDate.strftime("%d/%m/%y"))
					elif headers[col-1]=="Post end date":
						cell=ws["{}{}".format(colAlphabet,row)]
						cell.alignment=bottomRightVertical
						c=ws.cell(column=col,row=row,value=data.postEndDate.strftime("%d/%m/%y"))
					# elif  headers[col-1]== "Sum of Budget RIR":
					# 	cell=ws["{}{}".format(colAlphabet,row)]
					# 	cell.alignment=bottomRightVertical
					# 	c=ws.cell(column=col,row=row,value=data.budgetRir)
					# 	if data.budgetRir != None:
					# 		sumRir=sumRir+int(float(data.budgetRir))
			# ws['A{}'.format(str(rirSumLine))]="Grand Total"
			# ws['E{}'.format(str(rirSumLine))]=str(sumRir)
			ws['A{}'.format(str(rirSumLine))].font=boldFont
			ws['E{}'.format(str(rirSumLine))].font=boldFont
			ws['E{}'.format(str(rirSumLine))].alignment=bottomRightVertical
			ws['A{}'.format(str(rirSumLine))].fill=fillCellStyle
			ws['B{}'.format(str(rirSumLine))].fill=fillCellStyle
			ws['C{}'.format(str(rirSumLine))].fill=fillCellStyle
			ws['D{}'.format(str(rirSumLine))].fill=fillCellStyle
			# ws['E{}'.format(str(rirSumLine))].fill=fillCellStyle
			ws.merge_cells("A{}:C{}".format(str(rirSumLine+3),str(rirSumLine+3)))
			ws.merge_cells("A{}:C{}".format(str(rirSumLine+4),str(rirSumLine+4)))
			ws['A{}'.format(str(rirSumLine+3))]="Mong Quý Khách Hàng cùng hợp tác với Công ty để đảm bảo các chương trình thực thi hiệu quả trong thời gian tới."
			ws['A{}'.format(str(rirSumLine+4))]="Nếu Quý Khách Hàng có bất kỳ vấn đề nào cần làm rõ, vui lòng cho KCV được biết để cùng trao đổi"
			ws['A{}'.format(str(rirSumLine+6))]="Trân trọng cảm ơn Quý Khách Hàng"
			ws['A{}'.format(str(rirSumLine+7))]="Trưởng bộ phận quản lý kênh hiện đại"
			
			fileName="{}{}_{}.xlsx".format(f,loaict,str(datetime.now().date()))
			ws.print_area = 'A1:D{}'.format(str(rirSumLine+7))
			# Printer Settings
			ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
			ws.page_setup.paperSize = ws.PAPERSIZE_A4
			wb.save(fileName)
			sw.App.visible = False
			xl = sw.Book(fileName)
			xl.sheets("Thư thông báo").to_pdf(path=r'{}\\PDFs\\{}'.format( os.getcwdb().decode('utf-8'),fileName.replace(".xlsx","")))
			xl.close()
			pdf=open(r'{}\\PDFs\\{}'.format( os.getcwdb().decode('utf-8'),fileName.replace(".xlsx",".pdf")), "rb")
			os.remove(os.getcwdb().decode('utf-8') + "\\{}".format(fileName))
			pdffile=pdfFile()
			pdffile.masterFile=file
			pdffile.slaveFile.save(fileName.replace(".xlsx",".pdf"),File(pdf))	
			pdf.close()
			os.remove(os.getcwdb().decode('utf-8')+"\\PDFs\\{}.pdf".format(fileName.replace(".xlsx","")))
	elif loaiAccount!="All" and loaict=="All":
		
		wb=	openpyxl.Workbook()
		ws = wb.active
		ws.title="Thư Thông Báo"
		rirSumLine=0

		
		ws.merge_cells("A8:C8")
		
		# FORMAT COLUMN'S WIDTH
		ws.column_dimensions['A'].width=60
		ws.column_dimensions['B'].width=40
		ws.column_dimensions['C'].width=14
		ws.column_dimensions['D'].width=14
		# INSERT IMAGE
		img=Image("image\kimberlylogo.png")
		
		img.width=270
		img.height=30
		ws.add_image(img,"A1")	
		#TITLE,"Tp.HCM, Ngày","Kính gửi : Quý Khách Hàng Kênh Hiện Đại",
		# "Công ty TNHH Kimberly-Clark Việt Nam (Công ty)  xin trân trọng thông báo chương trình đến Quý Khách Hàng như thông tin đính kèm,"
		#"Loại CT","Account"
		ws["B4"]="THÔNG BÁO VỀ CHƯƠNG TRÌNH KHUYẾN MÃI" 
		ws["C5"]="Tp.HCM, Ngày"
		ws['D5']=str(datetime.now().date().strftime("%d/%m/%y"))
		ws["A6"]="Kính gửi : Quý Khách Hàng Kênh Hiện Đại" 
		ws['A8']="Công ty TNHH Kimberly-Clark Việt Nam (Công ty)  xin trân trọng thông báo chương trình đến Quý Khách Hàng như thông tin đính kèm,"
		ws['A10']="Loại CT"
		ws['B10']="All"
		ws['B11']=loaiAccount
		ws['A11']="Account"
		#CHANGE  FONT STYLES
		b4=ws["B4"];	c5=ws["C5"];	d5=ws['D5']
		a6=ws["A6"];	a8=ws['A8'];	a10=ws['A10']
		a11=ws['A11'];	b10=ws['B10'];	b11=ws['B11']
		b4.font=boldFont;				a6.font=boldFont; 		c5.font=italicFont;				
		d5.font=italicFont;				a8.font=normalFont;		a10.font=normalFont;	
		a11.font=normalFont;			b10.font=normalFont;	b11.font=normalFont
		#FILL CELL COLOR
		a10.fill=fillCellStyle;		a11.fill=fillCellStyle;		b10.fill=fillCellStyle
		b11.fill=fillCellStyle
		#///////////////////////////////////////////////////////////

		#header

		headers=['Mechanics: get/discount',"Product","Post start date","Post end date"]
		for header in range(0,len(headers)):
			_=ws.cell(column=header+1,row=13,value=headers[header])
			_.font=boldFont
			_.fill=fillCellStyle

		fileData= excel.objects.filter(filename=file,account=loaiAccount)
		
	elif loaiAccount!="All" and loaict!="All":
		print("4")
		listAcc=loaiAccount.split(',')
		listCt=loaict.split(',')
		for f in listAcc:
			
			title = 'THÔNG BÁO VỀ CHƯƠNG TRÌNH KHUYẾN MÃI' #.encode('utf-8')
			date_year = "Tp.HCM, Ngày {}".format(str(datetime.now().date().strftime("%d/%m/%Y")))
			pdf = PDF()
			pdf.add_font('Roboto','',r"static/Fonts/Roboto-Italic.ttf",uni=True)
			pdf.add_font('Roboto','B',r"static/Fonts/Roboto-BoldItalic.ttf",uni=True)
			pdf.add_font('Roboto','I',r"static/Fonts/Roboto-Italic.ttf",uni=True)
			pdf.set_font('Roboto','', size=10)
			# pdf=FPDF("L", "mm", "A4")

			pdf.add_page()
			title_w = pdf.get_string_width(title) + 6

			doc_w = pdf.w
			center = (doc_w - title_w) / 2
			#tittle
			pdf.text(center,15,title)
			#logo
			pdf.image('static/image/kimberlylogo.png', x = 2, w = 50,h=10, y=2)
			# date year

			pdf.set_font('Roboto','', size=8)
			pdf.text(165,20,date_year)
			# Kính gửi
			pdf.set_font('Roboto', 'B', size=9)
			text = 'Kính gửi: Quý Khách Hàng Kênh Hiện Đại'

			pdf.text(1, 25, text)
			# Tên chương trình
			pdf.set_font('Roboto','', size=8)
			text = "Công ty TNHH Kimberly-Clark Việt Nam (Công ty)  xin trân trọng thông báo chương trình đến Quý Khách Hàng như thông tin đính kèm,"

			pdf.text(1,32,text)
			# Loại CT

			#loai ct,loai acc
			print(pdf.get_x())
			print(pdf.get_y())
			pdf.set_y(pdf.get_y()+30)
			pdf.set_fill_color(153,204,255)
			pdf.cell(60,5,"Loại CT",0,0,"L",1)

			pdf.cell(60,5,"{}".format(loaict),0,1,"L",1)
			pdf.set_fill_color(153,204,255)
			pdf.set_fill_color(153,204,255)
			pdf.cell(60,5,"Account","B",0,"L",1)
			pdf.set_fill_color(153,204,255)
			pdf.cell(60,5,"{}".format(f),"B",1,"L",1)
			pdf.set_y(pdf.get_y()+5)
			#table header
			headers=['Mechanics: get/discount',"Product","Post start date","Post end date"]
			for i,header in enumerate(headers):
				if i==0:
					pdf.set_font('Roboto', 'B', size=9)
					pdf.set_fill_color(153,204,255)
					pdf.cell(60,5,"{}".format(headers[i]),"B",0,"L",1)
				elif i==1:
					pdf.set_font('Roboto', 'B', size=9)
					pdf.set_fill_color(153,204,255)
					pdf.cell(60,5,"{}".format(headers[i]),"B",0,"L",1)
				elif i ==2:
					pdf.set_font('Roboto', 'B', size=9)
					pdf.set_fill_color(153,204,255)
					pdf.cell(30,5,"{}".format(headers[i]),"B",0,"L",1)
				else:
					pdf.set_font('Roboto', 'B', size=9)
					pdf.set_fill_color(153,204,255)
					pdf.cell(30,5,"{}".format(headers[i]),"B",1,"L",1)

			#table data
			for ct in listCt:
				datas=excel.objects.filter(filename=file,account=f,loaiCt=ct)
				for row,data in enumerate(datas):
					for col,colAlphabet in enumerate(["A","B","C","D"]):
						mechanicsString=data.mechanicsGetORDiscount.replace("\n","")
						cellWitdhMax=60
						if pdf.get_string_width(mechanicsString) < cellWitdhMax:
							if headers[col]=='Mechanics: get/discount':
								string=data.mechanicsGetORDiscount
								pdf.cell(60,5,string,"B",0,"L")
							if headers[col]=='Product':
								string=data.product
								if string== None:
									pdf.cell(60,5,"","B",0,"L",1)
								else:
									pdf.cell(60,5,string,"B",0,"L")
							elif headers[col]=="Post start date":
								pdf.cell(30,5,"{}".format(data.postStartDate.strftime("%d/%m/%Y")),"B",0,"L")
							elif headers[col]=="Post end date":
								pdf.cell(30,5,"{}".format(data.postEndDate.strftime("%d/%m/%Y")),"B",1,"L")
						else:
							mechanicsStringLen=len(mechanicsString)
							startChar=0
							maxChar=0
							holdStringEachLine=[]
							holdStringTemp=""
							while (startChar<=mechanicsStringLen):
								while (pdf.get_string_width(holdStringTemp) < cellWitdhMax) and ((startChar+maxChar)<=mechanicsStringLen):
									maxChar=maxChar+1
									holdStringTemp= mechanicsString[startChar:(maxChar+startChar)]						
								startChar=startChar+maxChar
								holdStringEachLine.append(holdStringTemp)
								#reset
								maxChar=0
								holdStringTemp=""					
							print(holdStringEachLine)
							line=len(holdStringEachLine) #Numbers of line
							if headers[col]=='Mechanics: get/discount':
								
								xPos=pdf.get_x()
								yPos=pdf.get_y()
								pdf.multi_cell(cellWitdhMax,5,mechanicsString,"B")
								pdf.set_xy(xPos+cellWitdhMax,yPos)
							if headers[col]=='Product':
								string=data.product
								if string== None:
									pdf.cell(60,5*line,"","B",0,"L")
								else:
									pdf.cell(60,5*line,string,"B",0,"L",)
							elif headers[col]=="Post start date":
								pdf.cell(30,5*line,"{}".format(data.postStartDate.strftime("%d/%m/%Y")),"B",0,"L")
							elif headers[col]=="Post end date":
								pdf.cell(30,5*line,"{}".format(data.postEndDate.strftime("%d/%m/%Y")),"B",1,"L")

			#table footer
			for i,header in enumerate(headers):
				if i==0:
					pdf.set_font('Roboto', 'B', size=9)
					pdf.set_fill_color(153,204,255)
					pdf.cell(60,5,"","B",0,"L",1)
				elif i==1:
					pdf.set_font('Roboto', 'B', size=9)
					pdf.set_fill_color(153,204,255)
					pdf.cell(60,5,"","B",0,"L",1)
				elif i ==2:
					pdf.set_font('Roboto', 'B', size=9)
					pdf.set_fill_color(153,204,255)
					pdf.cell(30,5,"","B",0,"L",1)
				else:
					pdf.set_font('Roboto', 'B', size=9)
					pdf.set_fill_color(153,204,255)
					pdf.cell(30,5,"","B",1,"L",1)
					
			#add footer
			pdf.set_y(pdf.get_y()+30)
			pdf.cell(100,5,"Mong Quý Khách Hàng cùng hợp tác với Công ty để đảm bảo các chương trình thực thi hiệu quả trong thời gian tới.",0,1)
			pdf.cell(100,5,"Nếu Quý Khách Hàng có bất kỳ vấn đề nào cần làm rõ, vui lòng cho KCV được biết để cùng trao đổi",0,1)

			pdf.set_y(pdf.get_y()+20)
			pdf.cell(40,5,"Trân trọng cảm ơn Quý Khách Hàng",0,1)
			pdf.cell(40,5,"Trưởng bộ phận quản lý kênh hiện đại",0,1)

			pdf.set_y(pdf.get_y()+20)
			pdf.cell(40,5,"Phạm Nguyên Thủ",0,1)
			
			pdf.output("pdffile.pdf")
			# pdf.write_html(html)
			filename="{}{}_{}.pdf".format(f,loaict,str(datetime.now().date()))
			with open("pdffile.pdf",'rb') as pdf:
				pdffile=pdfFile()
				pdffile.masterFile=file
				pdffile.slaveFile.save(filename,File(pdf))	
	return "Success"


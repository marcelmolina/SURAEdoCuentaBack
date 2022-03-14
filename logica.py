import openpyxl as opyxl
from openpyxl.styles import PatternFill
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
import string
import time
from io import BytesIO
import cx_Oracle
import os
from dotenv import load_dotenv
from apoyo import getHeadColumnsBonos
from apoyo import getHeadColumnsComisones
from apoyo import getTipoSubBono
from apoyo import getTableNamesComisiones
from apoyo import getquery
from apoyo import getheaderpdf
import asyncio
import cx_Oracle_async

async def get_oracle_pool(app):
	try:
		load_dotenv()
		DB_HOST = os.getenv('DB_HOST')
		DB_PORT = os.getenv('DB_PORT')
		DB_SERVICE = os.getenv('DB_SERVICE')
		DB_USER = os.getenv('DB_USER')
		DB_PWD = os.getenv('DB_PWD')
		DB_SCHEMA = os.getenv('DB_SCHEMA')
		host = DB_HOST
		port = DB_PORT
		service_name = DB_SERVICE
		user = DB_USER
		password = DB_PWD
		schema = DB_SCHEMA
		sid = cx_Oracle_async.makedsn(host, port, service_name=service_name)
		try:
			pool = await cx_Oracle_async.create_pool(user,password,sid)
			return True,"",pool
		except Exception as ex:
			app.logger.error(ex)
			return False, 'Error en la conexion con la base de datos.', 0
	except Exception as ex:
		app.logger.error(ex)
		return False, 'Error en la configuracion de la base de datos.', 0

def getconexion(app):
	try:
		load_dotenv()
		DB_HOST = os.getenv('DB_HOST')
		DB_PORT = os.getenv('DB_PORT')
		DB_SERVICE = os.getenv('DB_SERVICE')
		DB_USER = os.getenv('DB_USER')
		DB_PWD = os.getenv('DB_PWD')
		DB_SCHEMA = os.getenv('DB_SCHEMA')
		host = DB_HOST
		port = DB_PORT
		service_name = DB_SERVICE
		user = DB_USER
		password = DB_PWD
		schema = DB_SCHEMA
		sid = cx_Oracle.makedsn(host, port, service_name=service_name)
		try:
			connection = cx_Oracle.connect(f"{user}/{password}@{host}:{port}/{service_name}")
			return True,"",connection
		except Exception as ex:
			app.logger.error(ex)
			return False, 'Error en la conexion con la base de datos.', 0
	except Exception as ex:
		app.logger.error(ex)
		return False, 'Error en la configuracion de la base de datos.', 0

def getperiodo(P_CLAVE,P_MES,P_ANIO,app):
	try:
		tabla = "TSIA_MOVAGEN"
		if P_CLAVE == 'P':
			tabla = "TSIA_MOVPROM"
		con_est, con_mssg, connection = getconexion(app)
		if not con_est:
			return False, con_mssg, 0, 0
		app.logger.info("Conectado a la base de datos.")
		entro = False
		P_Fefin = ""
		P_Feini = ""
		query1 = f"SELECT MAX(femovimi)-1 FROM {tabla} WHERE cdconc = 'SA' AND TO_CHAR(femovimi,'MMYYYY') = LPAD('{P_MES}',2,'0')||'{P_ANIO}'"
		cur1 = connection.cursor()
		cur1.execute(query1)
		for data in cur1:
			entro = True
			P_Fefin = data[0].strftime('%Y-%m-%d')
		if entro:
			entro = False
			query2 = f"SELECT MAX(femovimi) FROM {tabla} WHERE cdconc = 'SA' AND femovimi < TO_DATE('{P_Fefin}','YYYY-MM-DD')"
			cur2 = connection.cursor()
			cur2.execute(query2)
			for data in cur2:
				entro = True
				P_Feini = data[0].strftime('%Y-%m-%d')
		if not entro:
			return False,"Error en la busqueda de fechas.",0,0
		return True,"",P_Feini,P_Fefin
	except Exception as ex:
		app.logger.error(ex)
		return False, 'Error en la busqueda de fechas', 0,0

async def comisiones_xlsx(P_Clave,P_Feini,P_Fefin,P_COD,app):
	try:
		app.logger.info("Entrando a Estado de cuentas de Comisiones XLSX ("+P_COD+")")
		con_est, con_mssg, pool = await get_oracle_pool(app)
		if not con_est:
			return False, con_mssg, 0, 0, 0
		app.logger.info("Conectado a la base de datos.")
		has_agent = False
		app.logger.info("Iniciando carga de cursores")
		cursors = await tasks(app, 13, P_COD, 'COMISION', P_Clave, P_Feini, P_Fefin, pool)
		if 'Excepcion' in cursors:
			return False, 'Hubo un error obteniendo la data', 0, 0, 0
		app.logger.info("Cargados todos los cursores.")
		libro_nombre =P_Clave+"_" + P_Feini.replace("/", "")+"_"+ P_Fefin.replace("/", "")+'_comisiones.xlsx'
		plantilla = "plantilla_agentes.xlsx"
		if P_COD == 'P':
			plantilla = "plantilla_promotor.xlsx"
		wb = opyxl.load_workbook(plantilla)
		app.logger.info("Cargado archivo de plantilla xlsx.")
		ws = wb.worksheets[0]
		ws.cell(row=1, column=6).value = "ESTADOS DE CUENTA DE COMISIONES"
		ws.title = "Estado de Cuenta de Comisiones"
		has_agent = False
		app.logger.info("Leyendo cursor de cabecera")
		for row in cursors[0]:
			has_agent = True
			for i in range(0, len(row) - 4):
				ws.cell(row=4 + i, column=9).value = row[i]
			for i in range(len(row) - 4, len(row)):
				ws.cell(row=i, column=4).value = row[i]
		if not has_agent:
			app.logger.error("La cabecera volvio vacia.")
			return False,'Identificador no encontrado',0,0,0
		del cursors[0]
		f = 13  # principal gestor de filas del archivo
		greyFill = PatternFill(fill_type='solid', start_color='d9d9d9', end_color='d9d9d9')
		# NUEVO BLOQUE SECUENCIAL
		c_count = 1
		for cursor in cursors:
			app.logger.info(f"Leyendo cursor -> ({c_count})")
			fila_totales = [0, 0, 0, 0, 0]
			lista = getHeadColumnsComisones("excel", c_count)
			alphabet_string = string.ascii_uppercase
			alphabet_list = list(alphabet_string)
			ws.cell(row=f, column=1).value = getTableNamesComisiones(c_count)
			ws.cell(row=f, column=1).font = Font(name='Arial', size=9, bold=True)
			f += 1
			j = 0
			for item in lista:
				ws.cell(row=f, column=j + 1).value = item
				ws.cell(row=f, column=j + 1).fill = greyFill
				ws.cell(row=f, column=j + 1).font = Font(name='Arial', size=9, bold=True)
				ws.cell(row=f, column=j + 1).alignment = Alignment(horizontal="center", vertical="center")
				columna = alphabet_list[ws.cell(row=f, column=j + 1).column - 1]
				multiplicador = 2
				lista_columnas_esp = ['A', 'B', 'C', 'H']
				ancho = len(item)
				if len(item) <= 6:
					ancho *= 2
				if len(item) > 6 or columna == 'A':
					ancho *= 1.1
				if columna in lista_columnas_esp:
					ancho = 25
				if c_count == 5:
					ws.column_dimensions[columna].width = ancho
				j += 1
			j = 0
			k = 0
			f += 1
			has_data= False
			for row in cursor:
				for i in range(0, len(row)):
					if c_count not in [5, 6, 7, 8]:
						valor = row[i]
						if c_count in [1, 2]:
							if i != 0:
								valor = "{:,.2f}".format(valor)
						if c_count in [3, 4, 9, 10]:
							if i != 0:
								valor = "{:,.2f}".format(valor)
						if c_count in [11, 12]:
							if i in [1, 3, 6, 7, 8]:
								valor = "{:,.2f}".format(valor)
						ws.cell(row=f, column=i + 1).value = valor
						ws.cell(row=f, column=i + 1).alignment = Alignment(horizontal="center", vertical="center")
						if len(str(valor)) > 17:
							ws.cell(row=f, column=i + 1).font = Font(name='Arial', size=8)
							if len(str(valor)) > 25:
								ws.cell(row=f, column=i + 1).font = Font(name='Arial', size=7)
					else:
						valor = row[i]
						if i in [10, 11, 14, 15]:
							valor = "{:,.2f}".format(valor)
						if i == 10:
							fila_totales[0] += abs(row[i])
						if i == 11:
							fila_totales[1] += abs(row[i])
						if i == 13:
							fila_totales[2] += abs(row[i])
						if i == 14:
							fila_totales[3] += abs(row[i])
						if i == 15:
							fila_totales[4] += abs(row[i])
						has_data = True
						ws.cell(row=f, column=i + 1).value = valor
						ws.cell(row=f, column=i + 1).alignment = Alignment(horizontal="center", vertical="center")
						if len(str(valor)) > 17:
							ws.cell(row=f, column=i + 1).font = Font(name='Arial', size=8)
							if len(str(valor)) > 25:
								ws.cell(row=f, column=i + 1).font = Font(name='Arial', size=7)
				f += 1
			#cursor.close()
			if c_count in [5, 6, 7, 8] and has_data:
				fila_totales[0] = "{:,.2f}".format(fila_totales[0])
				fila_totales[1] = "{:,.2f}".format(fila_totales[1])
				fila_totales[2] = "{:,.2f}".format(fila_totales[2])
				fila_totales[3] = "{:,.2f}".format(fila_totales[3])
				fila_totales[4] = "{:,.2f}".format(fila_totales[4])
				ws.cell(row=f, column=10).value = "TOTAL"
				ws.cell(row=f, column=10).alignment = Alignment(horizontal="center", vertical="center")
				ws.cell(row=f, column=11).value = fila_totales[0]
				ws.cell(row=f, column=11).alignment = Alignment(horizontal="center", vertical="center")
				ws.cell(row=f, column=12).value = fila_totales[1]
				ws.cell(row=f, column=12).alignment = Alignment(horizontal="center", vertical="center")
				ws.cell(row=f, column=14).value = fila_totales[2]
				ws.cell(row=f, column=14).alignment = Alignment(horizontal="center", vertical="center")
				ws.cell(row=f, column=15).value = fila_totales[3]
				ws.cell(row=f, column=15).alignment = Alignment(horizontal="center", vertical="center")
				ws.cell(row=f, column=16).value = fila_totales[4]
				ws.cell(row=f, column=16).alignment = Alignment(horizontal="center", vertical="center")
				f += 1
			c_count += 1
		# fin de bloque
		virtual_wb = BytesIO()
		app.logger.info("Construyendo reporte xlsx.")
		wb.save(virtual_wb)
		return True,"",virtual_wb.getvalue(), wb.mime_type,libro_nombre
	except Exception as ex:
		app.logger.error(ex)
		return False, 'Error generando el reporte', 0, 0,0


async def dotask(app,i,connection,query):
	try:
		async with connection.cursor() as cursor:
			await cursor.execute(query)
			app.logger.info(f"Retorna data del cursor #{i}.")
			return await cursor.fetchall()
	except Exception as ex:
		app.logger.error(ex)
		return "Excepcion"

async def tasks(app,number,P_COD, tipo, P_Clave, P_Feini, P_Fefin,pool):
	try:
		async with pool.acquire() as connection:
			alltasks = []
			for i in range(number):
				alltasks.append(asyncio.create_task(dotask(app,i,connection,getquery(P_COD, tipo, i, P_Clave, P_Feini, P_Fefin))))
			cursors = await asyncio.gather(*alltasks)
			return cursors
	except Exception as ex:
		app.logger.error(ex)
		return "Excepcion"


async def comisiones_pdf(P_Clave,P_Feini,P_Fefin,P_COD,app):
	try:
		print("Estado de cuentas de Comisiones PDF")
		app.logger.info("Entrando a Estado de cuentas de Comisiones PDF (" + P_COD + ")")
		con_est, con_mssg, pool = await get_oracle_pool(app)
		if not con_est:
			return False, con_mssg, 0, 0, 0
		app.logger.info("Conectado a la base de datos.")
		has_agent = False
		app.logger.info("Iniciando carga de cursores")
		cursors = await tasks(app,13,P_COD, 'COMISION', P_Clave, P_Feini, P_Fefin,pool)
		if 'Excepcion' in cursors:
			return False, 'Hubo un error obteniendo la data', 0, 0, 0
		app.logger.info("Cargados todos los cursores.")
		libro_nombre =P_Clave+"_" + P_Feini.replace("/", "")+"_"+ P_Fefin.replace("/", "")+'_comisiones.pdf'

		virtual_wb = BytesIO()
		doc = SimpleDocTemplate(virtual_wb,pagesize=landscape((432*mm, 546*mm)))
		flowables = []
		app.logger.info("Leyendo cursor de cabecera")
		for row in cursors[0]:
			has_agent = True
			lista_aux = []
			for i in range(0, len(row)):
				lista_aux.append(row[i])
		if not has_agent:
			app.logger.error("La cabecera volvio vacia.")
			return False, 'Identificador no encontrado', 0, 0,0
		header_all = getheaderpdf(P_COD,lista_aux,'COMISIONES')
		grid = [('FONTNAME', (0, 0), (-1,0), 'Courier-Bold')]
		tbl = Table(header_all)
		tbl.setStyle(grid)
		flowables.append(tbl)
		del cursors[0]
		c_count = 1
		tblstyle = TableStyle(
			[('GRID', (0, 0), (0, 0), 0.25, colors.gray), ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
			 ('FONTSIZE', (0, 0), (0, 0), 7)])
		for cursor in cursors:
			app.logger.info(f"Leyendo cursor -> ({c_count})")
			lista = getHeadColumnsComisones("pdf", c_count)
			data_cursor = []
			taux = Table([("", getTableNamesComisiones(c_count), ""), ("", "", "")])
			taux.setStyle(grid)
			flowables.append(taux)
			data_cursor.append(lista)
			fila_totales = ["", "", "", "", "", "", "", "TOTAL", 0, 0, "", 0, 0, 0, "", ""]
			for row in cursor:
				lista_aux = []
				for i in range(0, len(row)):
					if c_count not in [5,6,7,8]:
						valor = row[i]
						if c_count in [1,2]:
							if i != 0:
								if valor < 0:
									valor = "(" + "{:,.2f}".format(abs(valor)) + ")"
								else:
									valor = "{:,.2f}".format(valor)
						if c_count in [3,4,9,10]:
							if i != 0:
								if valor < 0:
									valor = "(" + "{:,.2f}".format(abs(valor)) + ")"
								else:
									valor = "{:,.2f}".format(valor)
						if c_count in [11,12]:
							if i in [1,3,6,7,8]:
								if valor < 0:
									valor = "(" + "{:,.2f}".format(abs(valor)) + ")"
								else:
									valor = "{:,.2f}".format(valor)
						lista_aux.append(valor)
					else:
						if i not in [1,6,17]:
							valor = row[i]
							if i in [10, 11, 14, 15]:
								if valor < 0:
									valor = "(" + "{:,.2f}".format(abs(valor)) + ")"
								else:
									valor = "{:,.2f}".format(valor)
							lista_aux.append(valor)
							if i == 10:
								fila_totales[8] += abs(row[i])
							if i == 11:
								fila_totales[9] += abs(row[i])
							if i == 13:
								fila_totales[11] += abs(row[i])
							if i == 14:
								fila_totales[12] += abs(row[i])
							if i == 15:
								fila_totales[13] += abs(row[i])
				data_cursor.append(lista_aux)
			if c_count in [5,6,7,8] and len(data_cursor)>0:
				fila_totales[8] = "{:,.2f}".format(fila_totales[8])
				fila_totales[9] = "{:,.2f}".format(fila_totales[9])
				fila_totales[11] = "{:,.2f}".format(fila_totales[11])
				fila_totales[12] = "{:,.2f}".format(fila_totales[12])
				fila_totales[13] = "{:,.2f}".format(fila_totales[13])
				data_cursor.append(fila_totales)
			tbl = Table(data_cursor)
			tbl.setStyle(tblstyle)
			flowables.append(tbl)
			flowables.append(Table([("", " ", ""), ("", "", "")]))
			c_count += 1
			#cursor.close()
		# fin de bloque
		app.logger.info("Construyendo reporte pdf.")
		doc.build(flowables,canvasmaker=PageNumCanvas)
		return True,"",virtual_wb.getvalue(),"application/pdf", libro_nombre
	except Exception as ex:
		app.logger.error(ex)
		return False, 'Error en la generacion del reporte.', 0, 0,0


async def bonos_pdf(P_Clave,P_Feini,P_Fefin,P_COD,app):
	try:
		app.logger.info("Entrando a Estado de cuentas de Bonos PDF (" + P_COD + ")")
		con_est, con_mssg, pool = await get_oracle_pool(app)
		if not con_est:
			return False, con_mssg, 0, 0, 0
		app.logger.info("Conectado a la base de datos.")
		has_agent = False
		app.logger.info("Iniciando carga de cursores")
		cursors = await tasks(app, 2, P_COD, 'BONO', P_Clave, P_Feini, P_Fefin, pool)
		if 'Excepcion' in cursors:
			return False, 'Hubo un error obteniendo la data', 0, 0, 0
		app.logger.info("Cargados todos los cursores.")
		virtual_wb = BytesIO()
		doc = SimpleDocTemplate(virtual_wb,pagesize=landscape((432*mm, 546*mm)))
		flowables = []
		libro_nombre = P_Clave+"_" + P_Feini.replace("/", "")+"_"+ P_Fefin.replace("/", "")+'_bonos.pdf'
		app.logger.info("Leyendo cursor de cabecera")
		for row in cursors[0]:
			has_agent = True
			lista_aux = []
			for i in range(0, len(row)):
				lista_aux.append(row[i])
		#cursors[0].close
		if not has_agent:
			app.logger.error("La cabecera volvio vacia.")
			return False, 'Identificador no encontrado', 0, 0,0
		header_all = getheaderpdf(P_COD,lista_aux,'BONOS')
		grid = [('FONTNAME', (0, 0), (-1,0), 'Courier-Bold')]
		tbl = Table(header_all)
		tbl.setStyle(grid)
		flowables.append(tbl)

		j = 0
		lista = getHeadColumnsBonos("PDF",P_COD)
		data_body = []
		lista_aux = []
		has_data=False
		for item in lista:
			lista_aux.append(item)
		data_body.append(lista_aux)
		app.logger.info("Leyendo cursor -> (1)")
		for row in cursors[1]:
			has_data=True
			lista_aux = []
			for i in range(0, len(row)):
				if P_COD in ["A", "P"]:
					if i < 21 and i not in [3,8,9,19]:
						if i == 2:
							lista_aux.append(getTipoSubBono(row[i]))
						else:
							lista_aux.append(row[i])
				if P_COD == 'UDI':
					if i not in [1,2,16]:
						lista_aux.append(row[i])
			data_body.append(lista_aux)
		#cursors[1].close()
		if not has_data:
			app.logger.error("La tabla de detalle volvio vacia.")
			return False, 'Error generando el reporte.', 0, 0,0
		tbl = Table(data_body)
		tblstyle = TableStyle([('GRID',(0,0),(0,0),0.25,colors.gray),('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),('FONTSIZE', (0, 0), (0, 0), 7)])
		tbl.setStyle(tblstyle)
		flowables.append(tbl)

		app.logger.info("Construyendo reporte pdf.")
		doc.build(flowables,canvasmaker=PageNumCanvas)
		return True,"",virtual_wb.getvalue(),"application/pdf", libro_nombre
	except Exception as ex:
		return False, 'Error generando el reporte.', 0, 0,0


async def bonos_xlx(P_Clave,P_Feini,P_Fefin,P_COD,app):
	try:
		app.logger.info("Entrando a Estado de cuentas de Bonos XLSX (" + P_COD + ")")
		con_est, con_mssg, pool = await get_oracle_pool(app)
		if not con_est:
			return False, con_mssg, 0, 0, 0
		app.logger.info("Conectado a la base de datos.")
		has_agent = False
		app.logger.info("Iniciando carga de cursores")
		cursors = await tasks(app, 2, P_COD, 'BONO', P_Clave, P_Feini, P_Fefin, pool)
		if 'Excepcion' in cursors:
			return False, 'Hubo un error obteniendo la data', 0, 0, 0
		app.logger.info("Cargados todos los cursores.")
		plantilla = "plantilla_agentes.xlsx"
		if P_COD == 'P':
			plantilla = "plantilla_promotor.xlsx"
		if P_COD == 'UDI':
			plantilla = "plantilla_udi.xlsx"
		wb = opyxl.load_workbook(plantilla)
		app.logger.info("Cargado archivo de plantilla xlsx.")
		ws = wb.worksheets[0]

		libro_nombre = P_Clave+"_" + P_Feini.replace("/", "")+"_"+ P_Fefin.replace("/", "")+'_bonos.xlsx'
		app.logger.info("Leyendo cursor de cabecera")
		for row in cursors[0]:
			has_agent = True
			for i in range(0, len(row) - 4):
				ws.cell(row=4 + i, column=9).value = row[i]
			for i in range(len(row) - 4, len(row)):
				ws.cell(row=i, column=4).value = row[i]
		#cursors[0].close()
		if not has_agent:
			app.logger.error("La cabecera volvio vacia.")
			return False, 'Identificador no encontrado.', 0, 0,0
		j = 0
		greyFill = PatternFill(fill_type='solid', start_color='d9d9d9', end_color='d9d9d9')
		lista = getHeadColumnsBonos("excel",P_COD)
		alphabet_string = string.ascii_uppercase
		alphabet_list = list(alphabet_string)
		for item in lista:
			ws.cell(row=13, column=j + 1).value = item
			ws.cell(row=13, column=j + 1).fill = greyFill
			ws.cell(row=13, column=j + 1).font = Font(name='Arial', size=9, bold=True)
			ws.cell(row=13, column=j + 1).alignment = Alignment(horizontal="center", vertical="center")
			columna = alphabet_list[ws.cell(row=13, column=j + 1).column - 1]
			multiplicador = 2
			lista_columnas_esp = ['B', 'G']
			ancho = len(item)
			if len(item) <= 6:
				ancho *= 2
			if len(item) > 6 or columna == 'A':
				ancho *= 1.1
			if columna in lista_columnas_esp:
				ancho = 25
			ws.column_dimensions[columna].width = ancho
			j += 1
		j = 0
		k = 0
		has_data = False
		app.logger.info("Leyendo cursor -> (1)")
		for row in cursors[1]:
			has_data = True
			lista_razones = []
			if P_COD in ["A","P"]:
				for i in range(0, len(row)):

					if i < 21:
						aux = 0
						if i >= 9:
							aux = 1
						valor = row[i]
						if i == 2:
							valor = getTipoSubBono(row[i])
						ws.cell(row=14 + j, column=i + 1 + aux).value = valor
						ws.cell(row=14 + j, column=i + 1 + aux).alignment = Alignment(horizontal="center",
																					  vertical="center")
						if len(str(valor)) > 17:
							ws.cell(row=14 + j, column=i + 1 + aux).font = Font(name='Arial', size=8)
							if len(str(valor)) > 25:
								ws.cell(row=14 + j, column=i + 1 + aux).font = Font(name='Arial', size=7)
					else:
						if i == 21:
							if row[i] == 'NO':
								ws.cell(row=14 + j, column=10).value = 'Sí'
							else:
								ws.cell(row=14 + j, column=10).value = 'No'
						if i > 21 and row[21] == 'SI':
							if row[i] is not None and len(row[i]) > 0:
								lista_razones.append(row[i])
				if row[21] == 'SI' and len(lista_razones) > 0:
					ws.cell(row=14 + j, column=11).value = ", ".join(lista_razones)
				if row[21] == 'NO':
					ws.cell(row=14 + j, column=11).value = ' '
			if P_COD == 'UDI':
				for i in range(0, len(row)):
					valor = row[i]
					if i == 2:
						valor = getTipoSubBono(row[i])
					ws.cell(row=14 + j, column=i + 1 ).value = valor
					ws.cell(row=14 + j, column=i + 1 ).alignment = Alignment(horizontal="center",vertical="center")
					if len(str(valor)) > 17:
						ws.cell(row=14 + j, column=i + 1 ).font = Font(name='Arial', size=8)
						if len(str(valor)) > 25:
							ws.cell(row=14 + j, column=i + 1 ).font = Font(name='Arial', size=7)
			j += 1
		#cursors[1].close()
		if not has_data:
			app.logger.error("La tabla de detalle volvio vacia.")
			return False, 'Error en la generacion del reporte.', 0, 0,0
		virtual_wb = BytesIO()
		app.logger.info("Construyendo reporte xlsx.")
		wb.save(virtual_wb)
		return True,"",virtual_wb.getvalue(),wb.mime_type,libro_nombre
	except Exception as ex:
		app.logger.error(ex)
		return False, 'Error en la generacion del reporte.', 0, 0,0


class PageNumCanvas(canvas.Canvas):
	"""
    http://code.activestate.com/recipes/546511-page-x-of-y-with-reportlab/
    http://code.activestate.com/recipes/576832/
    """

	# ----------------------------------------------------------------------
	def __init__(self, *args, **kwargs):
		"""Constructor"""
		canvas.Canvas.__init__(self, *args, **kwargs)
		self.pages = []

	# ----------------------------------------------------------------------
	def showPage(self):
		"""
        On a page break, add information to the list
        """
		self.pages.append(dict(self.__dict__))
		self._startPage()

	# ----------------------------------------------------------------------
	def save(self):
		"""
        Add the page number to each page (page x of y)
        """
		page_count = len(self.pages)

		for page in self.pages:
			self.__dict__.update(page)
			self.draw_page_number(page_count)
			canvas.Canvas.showPage(self)

		canvas.Canvas.save(self)

	# ----------------------------------------------------------------------
	def draw_page_number(self, page_count):
		"""
        Add the page number
        """
		page = "Pagina %s de %s" % (self._pageNumber, page_count)
		self.setFont("Helvetica", 11)
		self.drawRightString(480 * mm, 420 * mm, page)
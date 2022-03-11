# -*- coding: utf-8 -*-
import os
import string
import time
import cx_Oracle
import openpyxl as opyxl
from openpyxl.styles import PatternFill
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from flask import Flask, jsonify, request, Response, make_response
from io import BytesIO
import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
import os
from dotenv import load_dotenv

load_dotenv()
SERVER_PORT = os.getenv('SERVER_PORT')
DB_HOST = os.getenv('DB_HOST')
DB_PORT = os.getenv('DB_PORT')
DB_SERVICE = os.getenv('DB_SERVICE')
DB_USER = os.getenv('DB_USER')
DB_PWD = os.getenv('DB_PWD')
DB_SCHEMA = os.getenv('DB_SCHEMA')

app = Flask(__name__)
app.config['JSON_SORT_KEYS'] = False
context_path = "/api/estados-cuenta/"

# Testing Route
@app.route(context_path+'/ping', methods=['GET'])
def ping():
    return jsonify({'response': 'pong!'})


@app.route(context_path + '/agentes/bonos/excel', methods=['GET'])
def bono_agente_xlsx():
	try:
		print("""

					\  \      __             __             _                  /  /
					 >  >    |_  _| _       /  _|_ _       |_) _ __  _  _     <  < 
					/  /     |__(_|(_) o    \__ |_(_| o    |_)(_)| |(_)_>      \  \ 

					""")

		P_Clave = request.args['codigo']
		P_Feini = request.args['desde']
		P_Fefin = request.args['hasta']
		P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
		P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
		host = DB_HOST
		port = DB_PORT
		service_name = DB_SERVICE
		user = DB_USER
		password = DB_PWD
		schema = DB_SCHEMA
		sid = cx_Oracle.makedsn(host, port, service_name=service_name)
		# Declaracion de cursores a utilizar
		try:
			connection = cx_Oracle.connect(f"{user}/{password}@{host}:{port}/{service_name}")
			try:
				wb = opyxl.load_workbook("plantilla_agentes.xlsx")
				ws = wb.worksheets[0]
				statement = connection.cursor()
				c_head = connection.cursor()
				has_agent = False
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.HEADER_BONO_AGENTE ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_head  ); end;",
					c_head=c_head, Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				libro_nombre = P_Clave+"_" + P_Feini.replace("/", "")+"_"+ P_Fefin.replace("/", "")+'_bonos.xlsx'
				for row in c_head:
					has_agent = True
					for i in range(0, len(row) - 4):
						ws.cell(row=4 + i, column=9).value = row[i]
					for i in range(len(row) - 4, len(row)):
						ws.cell(row=i, column=4).value = row[i]

				c_body = connection.cursor()
				if not has_agent:
					statement.close()
					c_head.close()
					return make_response(jsonify(succes=False, message="Codigo de agente no encontrado"), 400)
				c_head.close()
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.BODY_BONO_AGENTE ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_body  ); end;",
					c_body=c_body, Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				j = 0
				greyFill = PatternFill(fill_type='solid', start_color='d9d9d9', end_color='d9d9d9')
				lista = getHeadColumns("excel")
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
				for row in c_body:
					has_data = True
					lista_razones = []
					for i in range(0, len(row)):

						if i < 21:
							aux = 0
							if i >=9:
								aux=1
							valor = row[i]
							if i == 2:
								valor = getTipoSubBono(row[i])
							ws.cell(row=14 + j, column=i + 1+aux).value = valor
							ws.cell(row=14 + j, column=i + 1+aux).alignment = Alignment(horizontal="center",vertical="center")
							if len(str(valor)) > 17:
								ws.cell(row=14 + j, column=i + 1+aux).font = Font(name='Arial', size=8)
								if len(str(valor)) > 25:
									ws.cell(row=14 + j, column=i + 1+aux).font = Font(name='Arial', size=7)
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
					j += 1
				statement.close()
				c_body.close()
				if not has_data:
					statement.close()
					c_body.close()
					return make_response(jsonify(succes=False, message="Codigo de agente no retorna data en esas fechas"), 400)
			except Exception as ex:
				app.logger.error(ex)
				statement.close()
				c_head.close()
		except Exception as ex:
			app.logger.error(ex)
			statement.close()
			c_head.close()
		print("\nTermina Proceso " + time.strftime("%X"))

		virtual_wb = BytesIO()
		wb.save(virtual_wb)
		return Response(virtual_wb.getvalue(), mimetype=wb.mime_type,headers={"Content-Disposition": "attachment;filename="+libro_nombre})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="Error en la generacion del archivo"),400)

@app.route(context_path + '/agentes/bonos/pdf', methods=['GET'])
def bono_agente_pdf():
	try:
		print("""

					\  \      __             __             _                  /  /
					 >  >    |_  _| _       /  _|_ _       |_) _ __  _  _     <  < 
					/  /     |__(_|(_) o    \__ |_(_| o    |_)(_)| |(_)_>      \  \ 

					""")

		P_Clave = request.args['codigo']
		P_Feini = request.args['desde']
		P_Fefin = request.args['hasta']
		P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
		P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
		host = DB_HOST
		port = DB_PORT
		service_name = DB_SERVICE
		user = DB_USER
		password = DB_PWD
		schema = DB_SCHEMA
		sid = cx_Oracle.makedsn(host, port, service_name=service_name)
		# Declaracion de cursores a utilizar
		try:
			connection = cx_Oracle.connect(f"{user}/{password}@{host}:{port}/{service_name}")
			try:
				has_agent = False
				statement = connection.cursor()
				c_head = connection.cursor()
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.HEADER_BONO_AGENTE ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_head  ); end;",
					c_head=c_head, Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				libro_nombre =P_Clave+"_" + P_Feini.replace("/", "")+"_"+ P_Fefin.replace("/", "")+'_bonos.pdf'

				virtual_wb = BytesIO()
				doc = SimpleDocTemplate(virtual_wb,pagesize=landscape((432*mm, 546*mm)))
				flowables = []


				data_header = []
				for row in c_head:
					has_agent = True
					lista_aux = []
					for i in range(0, len(row)):
						lista_aux.append(row[i])
				if not has_agent:
					statement.close()
					c_head.close()
					return make_response(jsonify(succes=False, message="Codigo de agente no encontrado"), 400)
				header_all = [("   ","   ","ESTADO DE CUENTA DE BONOS","   ","   "),
							  ("Nombre del SAT para Sura:", "Seguros SURA S.A. de C.V","     ","Nombre del Agente:",lista_aux[0]),
							  ("RFC de Sura:", "R.F.C R&S-811221KR6","   ","Clave del Agente:",lista_aux[1]),
							  ("Domicilio Sura:", "Blvd. Adolfo López Mateos No. 2448","   ","RFC del Agente:",lista_aux[2]),
							  ("Col. Altavista C.P. 01060 Ciudad de México.", "Tel 5723-7999","   ","Domicilio del Agente:",lista_aux[3]),
							  ("Periodo de corte:",lista_aux[8],"   ","Tipo de productor:",lista_aux[4]),
							  ("Desde:",lista_aux[9],"   ","Promotoría a la que pertenece:",lista_aux[5]),
							  ("Hasta:",lista_aux[10],"   ","Clave de productor:",lista_aux[6]),
							  ("Fecha de Preliquidación:",lista_aux[11],"   ","Cuenta bancaria dada de alta:",lista_aux[7]),
							  ("   ","   ","   ","   ","   ")]
				grid = [('FONTNAME', (0, 0), (0, -1), 'Courier-Bold'),
						('FONTNAME', (0, 0), (-1,0), 'Courier-Bold')]
				tbl = Table(header_all)
				tbl.setStyle(grid)
				flowables.append(tbl)
				c_head.close()
				c_body = connection.cursor()
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.BODY_BONO_AGENTE ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_body  ); end;",
					c_body=c_body, Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				j = 0
				lista = getHeadColumns("pdf")
				data_body = []
				lista_aux = []
				has_data=False
				for item in lista:
					lista_aux.append(item)
				data_body.append(lista_aux)
				for row in c_body:
					has_data=True
					lista_aux = []
					for i in range(0, len(row)):
						if i < 21 and i not in [3,8,9,19]:
							if i == 2:
								lista_aux.append(getTipoSubBono(row[i]))
							else:
								lista_aux.append(row[i])
					data_body.append(lista_aux)
				if not has_data:
					statement.close()
					c_body.close()
					return make_response(jsonify(succes=False, message="Codigo de agente no retorna data en esas fechas"), 400)
				tbl = Table(data_body)
				tblstyle = TableStyle([('GRID',(0,0),(0,0),0.25,colors.gray),('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),('FONTSIZE', (0, 0), (0, 0), 7)])
				tbl.setStyle(tblstyle)
				flowables.append(tbl)

				statement.close()
				c_body.close()
				doc.build(flowables)

			except Exception as ex:
				app.logger.error(ex)
				statement.close()
				c_head.close()
		except Exception as ex:
			app.logger.error(ex)
			statement.close()
			c_head.close()
		print("\nTermina Proceso " + time.strftime("%X"))

		return Response(virtual_wb.getvalue(), mimetype="application/pdf",
						 headers={"Content-Disposition": "attachment;filename=" + libro_nombre})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="Error en la generacion del archivo"),400)

@app.route(context_path + '/agentes/comisiones/pdf', methods=['GET'])
def comisiones_agente_pdf():
	try:
		print("Estado de cuentas comisiones")

		P_Clave = request.args['codigo']
		P_Feini = request.args['desde']
		P_Fefin = request.args['hasta']
		P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
		P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
		host = DB_HOST
		port = DB_PORT
		service_name = DB_SERVICE
		user = DB_USER
		password = DB_PWD
		schema = DB_SCHEMA
		sid = cx_Oracle.makedsn(host, port, service_name=service_name)
		# Declaracion de cursores a utilizar
		try:
			connection = cx_Oracle.connect(f"{user}/{password}@{host}:{port}/{service_name}")
			try:
				has_agent = False
				statement = connection.cursor()
				c_head = connection.cursor()
				c1 = connection.cursor()
				c2 = connection.cursor()
				c3 = connection.cursor()
				c4 = connection.cursor()
				c5 = connection.cursor()
				c6 = connection.cursor()
				c7 = connection.cursor()
				c8 = connection.cursor()
				c9 = connection.cursor()
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.BODY_COMISIONES_AGENTE ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_head,:c1,:c2,:c3,:c4,:c5,:c6,:c7,:c8,:c9); end;",
					c_head=c_head,c1=c1,c2=c2,c3=c3,c4=c4,c5=c5,c6=c6,c7=c7,c8=c8,c9=c9, Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				libro_nombre =P_Clave+"_" + P_Feini.replace("/", "")+"_"+ P_Fefin.replace("/", "")+'_comisiones.pdf'
				cursores = []
				cursores.append(c1)
				cursores.append(c2)
				cursores.append(c3)
				cursores.append(c4)
				cursores.append(c5)
				cursores.append(c6)
				cursores.append(c7)
				cursores.append(c8)
				cursores.append(c9)
				virtual_wb = BytesIO()
				doc = SimpleDocTemplate(virtual_wb,pagesize=landscape((432*mm, 546*mm)))
				flowables = []
				for row in c_head:
					has_agent = True
					lista_aux = []
					for i in range(0, len(row)):
						lista_aux.append(row[i])
				if not has_agent:
					statement.close()
					c_head.close()
					c1.close()
					c2.close()
					c3.close()
					c4.close()
					c5.close()
					c6.close()
					c7.close()
					c8.close()
					c9.close()
					return make_response(jsonify(succes=False, message="Codigo de agente no encontrado"), 400)
				header_all = [("   ","   ","ESTADO DE CUENTA DE COMISIONES","   ","   "),
							  ("Nombre del SAT para Sura:", "Seguros SURA S.A. de C.V","     ","Nombre del Agente:",lista_aux[0]),
							  ("RFC de Sura:", "R.F.C R&S-811221KR6","   ","Clave del Agente:",lista_aux[1]),
							  ("Domicilio Sura:", "Blvd. Adolfo López Mateos No. 2448","   ","RFC del Agente:",lista_aux[2]),
							  ("Col. Altavista C.P. 01060 Ciudad de México.", "Tel 5723-7999","   ","Domicilio del Agente:",lista_aux[3]),
							  ("Periodo de corte:",lista_aux[8],"   ","Tipo de productor:",lista_aux[4]),
							  ("Desde:",lista_aux[9],"   ","Promotoría a la que pertenece:",lista_aux[5]),
							  ("Hasta:",lista_aux[10],"   ","Clave de productor:",lista_aux[6]),
							  ("Fecha de Preliquidación:",lista_aux[11],"   ","Cuenta bancaria dada de alta:",lista_aux[7]),
							  ("   ","   ","   ","   ","   ")]
				grid = [('FONTNAME', (0, 0), (0, -1), 'Courier-Bold'),
						('FONTNAME', (0, 0), (-1,0), 'Courier-Bold')]
				tbl = Table(header_all)
				tbl.setStyle(grid)
				flowables.append(tbl)
				c_head.close()
				#se termino de generar el cabecero, ahor los cursores de detalle
				tblstyle = TableStyle(
					[('GRID', (0, 0), (0, 0), 0.25, colors.gray), ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
					 ('FONTSIZE', (0, 0), (0, 0), 7)])
				c_count=1
				for cursor in cursores:
					lista = getHeadColumnsComisones("pdf", c_count)
					data_cursor = []
					lista_aux = []
					taux = Table([("", getTableNamesComisiones(c_count), ""), ("", "", "")])
					taux.setStyle(grid)
					flowables.append(taux)
					data_cursor.append(lista)
					for row in cursor:
						lista_aux = []
						for i in range(0, len(row)):
							if c_count == 5:
								if i not in [1, 6, 17]:
									lista_aux.append(row[i])
							else:
								lista_aux.append(row[i])
						data_cursor.append(lista_aux)
					tbl = Table(data_cursor)
					tbl.setStyle(tblstyle)
					flowables.append(tbl)
					flowables.append(Table([("", " ", ""), ("", "", "")]))
					c_count += 1
					cursor.close()
				statement.close()

				doc.build(flowables)

			except Exception as ex:
				app.logger.error(ex)
				statement.close()
				c_head.close()
				c1.close()
				c2.close()
				c3.close()
				c4.close()
				c5.close()
				c6.close()
				c7.close()
				c8.close()
				c9.close()
		except Exception as ex:
			app.logger.error(ex)
			statement.close()
			c_head.close()
			c1.close()
			c2.close()
			c3.close()
			c4.close()
			c5.close()
			c6.close()
			c7.close()
			c8.close()
			c9.close()
		print("\nTermina Proceso " + time.strftime("%X"))

		return Response(virtual_wb.getvalue(), mimetype="application/pdf",
						 headers={"Content-Disposition": "attachment;filename=" + libro_nombre})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="Error en la generacion del archivo"),400)


@app.route(context_path + '/agentes/comisiones/excel', methods=['GET'])
def comisiones_agente_xlsx():
	try:
		print("Estado de cuentas de Comisiones")

		P_Clave = request.args['codigo']
		P_Feini = request.args['desde']
		P_Fefin = request.args['hasta']
		P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
		P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
		host = DB_HOST
		port = DB_PORT
		service_name = DB_SERVICE
		user = DB_USER
		password = DB_PWD
		schema = DB_SCHEMA
		sid = cx_Oracle.makedsn(host, port, service_name=service_name)
		# Declaracion de cursores a utilizar
		try:
			connection = cx_Oracle.connect(f"{user}/{password}@{host}:{port}/{service_name}")
			try:
				wb = opyxl.load_workbook("plantilla_agentes.xlsx")
				ws = wb.worksheets[0]
				ws.cell(row=1, column=6).value = "ESTADOS DE CUENTA DE COMISIONES"
				ws.title = "Estado de Cuenta de Comisiones"
				statement = connection.cursor()
				has_agent = False
				c_head = connection.cursor()
				c1 = connection.cursor()
				c2 = connection.cursor()
				c3 = connection.cursor()
				c4 = connection.cursor()
				c5 = connection.cursor()
				c6 = connection.cursor()
				c7 = connection.cursor()
				c8 = connection.cursor()
				c9 = connection.cursor()
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.BODY_COMISIONES_AGENTE ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_head,:c1,:c2,:c3,:c4,:c5,:c6,:c7,:c8,:c9); end;",
					c_head=c_head, c1=c1, c2=c2, c3=c3, c4=c4, c5=c5, c6=c6, c7=c7, c8=c8, c9=c9,
					Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				libro_nombre = P_Clave + "_" + P_Feini.replace("/", "") + "_" + P_Fefin.replace("/",
																								"") + '_comisiones.xlsx'
				for row in c_head:
					has_agent = True
					for i in range(0, len(row) - 4):
						ws.cell(row=4 + i, column=9).value = row[i]
					for i in range(len(row) - 4, len(row)):
						ws.cell(row=i, column=4).value = row[i]
				if not has_agent:
					statement.close()
					c_head.close()
					c1.close()
					c2.close()
					c3.close()
					c4.close()
					c5.close()
					c6.close()
					c7.close()
					c8.close()
					c9.close()
					return make_response(jsonify(succes=False, message="Codigo de agente no encontrado"), 400)
				c_head.close()
				f = 13 #principal gestor de filas del archivo
				greyFill = PatternFill(fill_type='solid', start_color='d9d9d9', end_color='d9d9d9')
				#NUEVO BLOQUE SECUENCIAL
				cursores = []
				cursores.append(c1)
				cursores.append(c2)
				cursores.append(c3)
				cursores.append(c4)
				cursores.append(c5)
				cursores.append(c6)
				cursores.append(c7)
				cursores.append(c8)
				cursores.append(c9)
				c_count = 1
				for cursor in cursores:
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
						lista_columnas_esp = ['A','B','C', 'H']
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
					for row in cursor:
						for i in range(0, len(row)):
							valor = row[i]
							ws.cell(row=f, column=i + 1).value = valor
							ws.cell(row=f, column=i + 1).alignment = Alignment(horizontal="center", vertical="center")
							if len(str(valor)) > 17:
								ws.cell(row=f, column=i + 1).font = Font(name='Arial', size=8)
								if len(str(valor)) > 25:
									ws.cell(row=f, column=i + 1).font = Font(name='Arial', size=7)
						f += 1
					cursor.close()
					c_count += 1
				statement.close()
			# fin de bloque
			except Exception as ex:
				app.logger.error(ex)
				statement.close()
				c_head.close()
				c1.close()
				c2.close()
				c3.close()
				c4.close()
				c5.close()
				c6.close()
				c7.close()
				c8.close()
				c9.close()
		except Exception as ex:
			app.logger.error(ex)
			statement.close()
			c_head.close()
			c1.close()
			c2.close()
			c3.close()
			c4.close()
			c5.close()
			c6.close()
			c7.close()
			c8.close()
			c9.close()
		print("\nTermina Proceso " + time.strftime("%X"))

		virtual_wb = BytesIO()
		wb.save(virtual_wb)
		return Response(virtual_wb.getvalue(), mimetype=wb.mime_type,headers={"Content-Disposition": "attachment;filename="+libro_nombre})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="Error en la generacion del archivo"),400)


@app.route(context_path + '/promotores/bonos/excel', methods=['GET'])
def bono_promotores_xlsx():
	try:
		print("""

					\  \      __             __             _                  /  /
					 >  >    |_  _| _       /  _|_ _       |_) _ __  _  _     <  < 
					/  /     |__(_|(_) o    \__ |_(_| o    |_)(_)| |(_)_>      \  \ 

					""")

		P_Clave = request.args['codigo']
		P_Feini = request.args['desde']
		P_Fefin = request.args['hasta']
		P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
		P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
		host = DB_HOST
		port = DB_PORT
		service_name = DB_SERVICE
		user = DB_USER
		password = DB_PWD
		schema = DB_SCHEMA
		sid = cx_Oracle.makedsn(host, port, service_name=service_name)
		# Declaracion de cursores a utilizar
		try:
			connection = cx_Oracle.connect(f"{user}/{password}@{host}:{port}/{service_name}")
			try:
				wb = opyxl.load_workbook("plantilla_promotor.xlsx")
				ws = wb.worksheets[0]
				statement = connection.cursor()
				c_head = connection.cursor()
				has_agent = False
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.HEADER_BONO_PROMOTOR ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_head  ); end;",
					c_head=c_head, Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				libro_nombre = P_Clave+"_" + P_Feini.replace("/", "")+"_"+ P_Fefin.replace("/", "")+'_bonos.xlsx'
				for row in c_head:
					has_agent = True
					for i in range(0, len(row) - 4):
						ws.cell(row=4 + i, column=9).value = row[i]
					for i in range(len(row) - 4, len(row)):
						ws.cell(row=i, column=4).value = row[i]

				c_body = connection.cursor()
				if not has_agent:
					statement.close()
					c_head.close()
					return make_response(jsonify(succes=False, message="Codigo de promotor no encontrado"), 400)
				c_head.close()
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.BODY_BONO_PROMOTOR ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_body  ); end;",
					c_body=c_body, Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				j = 0
				greyFill = PatternFill(fill_type='solid', start_color='d9d9d9', end_color='d9d9d9')
				lista = getHeadColumns("excel")
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
				for row in c_body:
					has_data = True
					lista_razones = []
					for i in range(0, len(row)):

						if i < 21:
							aux = 0
							if i >=9:
								aux=1
							valor = row[i]
							if i ==2:
								valor = getTipoSubBono(valor)
							ws.cell(row=14 + j, column=i + 1+aux).value = valor
							ws.cell(row=14 + j, column=i + 1+aux).alignment = Alignment(horizontal="center",vertical="center")
							if len(str(valor)) > 17:
								ws.cell(row=14 + j, column=i + 1+aux).font = Font(name='Arial', size=8)
								if len(str(valor)) > 25:
									ws.cell(row=14 + j, column=i + 1+aux).font = Font(name='Arial', size=7)
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
					j += 1
				statement.close()
				c_body.close()
				if not has_data:
					statement.close()
					c_body.close()
					return make_response(jsonify(succes=False, message="Codigo de promotor no retorna data en esas fechas"), 400)
			except Exception as ex:
				app.logger.error(ex)
				statement.close()
				c_head.close()
		except Exception as ex:
			app.logger.error(ex)
			statement.close()
			c_head.close()
		print("\nTermina Proceso " + time.strftime("%X"))

		virtual_wb = BytesIO()
		wb.save(virtual_wb)
		return Response(virtual_wb.getvalue(), mimetype=wb.mime_type,headers={"Content-Disposition": "attachment;filename="+libro_nombre})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="Error en la generacion del archivo"),400)

@app.route(context_path + '/promotores/bonos/pdf', methods=['GET'])
def bono_promotores_pdf():
	try:
		print("""

					\  \      __             __             _                  /  /
					 >  >    |_  _| _       /  _|_ _       |_) _ __  _  _     <  < 
					/  /     |__(_|(_) o    \__ |_(_| o    |_)(_)| |(_)_>      \  \ 

					""")

		P_Clave = request.args['codigo']
		P_Feini = request.args['desde']
		P_Fefin = request.args['hasta']
		P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
		P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
		host = DB_HOST
		port = DB_PORT
		service_name = DB_SERVICE
		user = DB_USER
		password = DB_PWD
		schema = DB_SCHEMA
		sid = cx_Oracle.makedsn(host, port, service_name=service_name)
		# Declaracion de cursores a utilizar
		try:
			connection = cx_Oracle.connect(f"{user}/{password}@{host}:{port}/{service_name}")
			try:
				has_agent = False
				statement = connection.cursor()
				c_head = connection.cursor()
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.HEADER_BONO_PROMOTOR ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_head  ); end;",
					c_head=c_head, Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				libro_nombre =P_Clave+"_" + P_Feini.replace("/", "")+"_"+ P_Fefin.replace("/", "")+'_bonos.pdf'

				virtual_wb = BytesIO()
				doc = SimpleDocTemplate(virtual_wb,pagesize=landscape((432*mm, 546*mm)))
				flowables = []


				data_header = []
				for row in c_head:
					has_agent = True
					lista_aux = []
					for i in range(0, len(row)):
						lista_aux.append(row[i])
				if not has_agent:
					statement.close()
					c_head.close()
					return make_response(jsonify(succes=False, message="Codigo de agente no encontrado"), 400)
				header_all = [("   ","   ","ESTADO DE CUENTA DE BONOS","   ","   "),
							  ("Nombre del SAT para Sura:", "Seguros SURA S.A. de C.V","     ","Nombre del Promotor:",lista_aux[0]),
							  ("RFC de Sura:", "R.F.C R&S-811221KR6","   ","Clave del Promotor:",lista_aux[1]),
							  ("Domicilio Sura:", "Blvd. Adolfo López Mateos No. 2448","   ","RFC del Promotor:",lista_aux[2]),
							  ("Col. Altavista C.P. 01060 Ciudad de México.", "Tel 5723-7999","   ","Domicilio del Promotor:",lista_aux[3]),
							  ("Periodo de corte:",lista_aux[8],"   ","Tipo de productor:",lista_aux[4]),
							  ("Desde:",lista_aux[9],"   ","Promotoría a la que pertenece:",lista_aux[5]),
							  ("Hasta:",lista_aux[10],"   ","Clave de productor:",lista_aux[6]),
							  ("Fecha de Preliquidación:",lista_aux[11],"   ","Número exterior:",lista_aux[7]),
							  ("   ","   ","   ","   ","   ")]
				grid = [('FONTNAME', (0, 0), (0, -1), 'Courier-Bold'),
						('FONTNAME', (0, 0), (-1,0), 'Courier-Bold')]
				tbl = Table(header_all)
				tbl.setStyle(grid)
				flowables.append(tbl)
				c_head.close()
				c_body = connection.cursor()
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.BODY_BONO_PROMOTOR ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_body  ); end;",
					c_body=c_body, Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				j = 0
				lista = getHeadColumns("pdf")
				data_body = []
				lista_aux = []
				has_data=False
				for item in lista:
					lista_aux.append(item)
				data_body.append(lista_aux)
				for row in c_body:
					has_data=True
					lista_aux = []
					for i in range(0, len(row)):
						if i < 21 and i not in [3,8,9,19]:
							if i == 2:
								lista_aux.append(getTipoSubBono(row[i]))
							else:
								lista_aux.append(row[i])
					data_body.append(lista_aux)
				if not has_data:
					statement.close()
					c_body.close()
					return make_response(jsonify(succes=False, message="Codigo de promotor no retorna data en esas fechas"), 400)
				tbl = Table(data_body)
				tblstyle = TableStyle([('GRID',(0,0),(0,0),0.25,colors.gray),('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),('FONTSIZE', (0, 0), (0, 0), 7)])
				tbl.setStyle(tblstyle)
				flowables.append(tbl)

				statement.close()
				c_body.close()
				doc.build(flowables)

			except Exception as ex:
				app.logger.error(ex)
				statement.close()
				c_head.close()
		except Exception as ex:
			app.logger.error(ex)
			statement.close()
			c_head.close()
		print("\nTermina Proceso " + time.strftime("%X"))

		return Response(virtual_wb.getvalue(), mimetype="application/pdf",
						 headers={"Content-Disposition": "attachment;filename=" + libro_nombre})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="Error en la generacion del archivo"),400)


@app.route(context_path + '/promotores/comisiones/pdf', methods=['GET'])
def comisiones_promotores_pdf():
	try:
		print("Estado de cuentas comisiones")

		P_Clave = request.args['codigo']
		P_Feini = request.args['desde']
		P_Fefin = request.args['hasta']
		P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
		P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
		host = DB_HOST
		port = DB_PORT
		service_name = DB_SERVICE
		user = DB_USER
		password = DB_PWD
		schema = DB_SCHEMA
		sid = cx_Oracle.makedsn(host, port, service_name=service_name)
		# Declaracion de cursores a utilizar
		try:
			connection = cx_Oracle.connect(f"{user}/{password}@{host}:{port}/{service_name}")
			try:
				has_agent = False
				statement = connection.cursor()
				c_head = connection.cursor()
				c1 = connection.cursor()
				c2 = connection.cursor()
				c3 = connection.cursor()
				c4 = connection.cursor()
				c5 = connection.cursor()
				c6 = connection.cursor()
				c7 = connection.cursor()
				c8 = connection.cursor()
				c9 = connection.cursor()
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.BODY_COMISIONES_PROMOTOR ( :Pb_AGENTE, :Pb_FEINI, :Pb_FEFIN, :c_head,:c1,:c2,:c3,:c4,:c5,:c6,:c7,:c8,:c9); end;",
					c_head=c_head,c1=c1,c2=c2,c3=c3,c4=c4,c5=c5,c6=c6,c7=c7,c8=c8,c9=c9, Pb_AGENTE=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				libro_nombre =P_Clave+"_" + P_Feini.replace("/", "")+"_"+ P_Fefin.replace("/", "")+'_comisiones.pdf'

				virtual_wb = BytesIO()
				doc = SimpleDocTemplate(virtual_wb,pagesize=landscape((432*mm, 546*mm)))
				flowables = []
				for row in c_head:
					has_agent = True
					lista_aux = []
					for i in range(0, len(row)):
						lista_aux.append(row[i])
				if not has_agent:
					statement.close()
					c_head.close()
					c1.close()
					c2.close()
					c3.close()
					c4.close()
					c5.close()
					c6.close()
					c7.close()
					c8.close()
					c9.close()
					return make_response(jsonify(succes=False, message="Codigo de agente no encontrado"), 400)
				header_all = [("   ","   ","ESTADO DE CUENTA DE COMISIONES","   ","   "),
							  ("Nombre del SAT para Sura:", "Seguros SURA S.A. de C.V","     ","Nombre del Promotor:",lista_aux[0]),
							  ("RFC de Sura:", "R.F.C R&S-811221KR6","   ","Clave del Promotor:",lista_aux[1]),
							  ("Domicilio Sura:", "Blvd. Adolfo López Mateos No. 2448","   ","RFC del Agente:",lista_aux[2]),
							  ("Col. Altavista C.P. 01060 Ciudad de México.", "Tel 5723-7999","   ","Domicilio del Promotor:",lista_aux[3]),
							  ("Periodo de corte:",lista_aux[8],"   ","Tipo de productor:",lista_aux[4]),
							  ("Desde:",lista_aux[9],"   ","Promotoría a la que pertenece:",lista_aux[5]),
							  ("Hasta:",lista_aux[10],"   ","Clave de productor:",lista_aux[6]),
							  ("Fecha de Preliquidación:",lista_aux[11],"   ","Número exterior:",lista_aux[7]),
							  ("   ","   ","   ","   ","   ")]
				grid = [('FONTNAME', (0, 0), (0, -1), 'Courier-Bold'),
						('FONTNAME', (0, 0), (-1,0), 'Courier-Bold')]
				tbl = Table(header_all)
				tbl.setStyle(grid)
				flowables.append(tbl)
				c_head.close()
				c_count = 1
				cursores = []
				cursores.append(c1)
				cursores.append(c2)
				cursores.append(c3)
				cursores.append(c4)
				cursores.append(c5)
				cursores.append(c6)
				cursores.append(c7)
				cursores.append(c8)
				cursores.append(c9)

				tblstyle = TableStyle(
					[('GRID', (0, 0), (0, 0), 0.25, colors.gray), ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
					 ('FONTSIZE', (0, 0), (0, 0), 7)])
				for cursor in cursores:
					lista = getHeadColumnsComisones("pdf", c_count)
					data_cursor = []
					lista_aux = []
					taux = Table([("", getTableNamesComisiones(c_count), ""), ("", "", "")])
					taux.setStyle(grid)
					flowables.append(taux)
					data_cursor.append(lista)
					for row in cursor:
						lista_aux = []
						for i in range(0, len(row)):
							lista_aux.append(row[i])
						data_cursor.append(lista_aux)
					tbl = Table(data_cursor)
					tbl.setStyle(tblstyle)
					flowables.append(tbl)
					flowables.append(Table([("", " ", ""), ("", "", "")]))
					c_count += 1
					cursor.close()

				# fin de bloque
				statement.close()

				doc.build(flowables)

			except Exception as ex:
				app.logger.error(ex)
				statement.close()
				c_head.close()
				c1.close()
				c2.close()
				c3.close()
				c4.close()
				c5.close()
				c6.close()
				c7.close()
				c8.close()
				c9.close()
		except Exception as ex:
			app.logger.error(ex)
			statement.close()
			c_head.close()
			c1.close()
			c2.close()
			c3.close()
			c4.close()
			c5.close()
			c6.close()
			c7.close()
			c8.close()
			c9.close()
		print("\nTermina Proceso " + time.strftime("%X"))

		return Response(virtual_wb.getvalue(), mimetype="application/pdf",
						 headers={"Content-Disposition": "attachment;filename=" + libro_nombre})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="Error en la generacion del archivo"),400)

@app.route(context_path + '/promotores/comisiones/excel', methods=['GET'])
def comisiones_promotor_xlsx():
	try:
		print("EStado de cuentas de Comisiones")

		P_Clave = request.args['codigo']
		P_Feini = request.args['desde']
		P_Fefin = request.args['hasta']
		P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
		P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
		host = DB_HOST
		port = DB_PORT
		service_name = DB_SERVICE
		user = DB_USER
		password = DB_PWD
		schema = DB_SCHEMA
		sid = cx_Oracle.makedsn(host, port, service_name=service_name)
		# Declaracion de cursores a utilizar
		try:
			connection = cx_Oracle.connect(f"{user}/{password}@{host}:{port}/{service_name}")
			try:
				wb = opyxl.load_workbook("plantilla_promotor.xlsx")
				ws = wb.worksheets[0]
				ws.cell(row=1, column=6).value = "ESTADOS DE CUENTA DE COMISIONES"
				ws.title = "Estado de Cuenta de Comisiones"
				statement = connection.cursor()
				has_agent = False
				c_head = connection.cursor()
				c1 = connection.cursor()
				c2 = connection.cursor()
				c3 = connection.cursor()
				c4 = connection.cursor()
				c5 = connection.cursor()
				c6 = connection.cursor()
				c7 = connection.cursor()
				c8 = connection.cursor()
				c9 = connection.cursor()
				statement.execute(
					"begin "+schema+".PKG_MUI_ESTADOS_DE_CUENTA_1.BODY_COMISIONES_PROMOTOR ( :Pb_PROMOTOR, :Pb_FEINI, :Pb_FEFIN, :c_head,:c1,:c2,:c3,:c4,:c5,:c6,:c7,:c8,:c9); end;",
					c_head=c_head, c1=c1, c2=c2, c3=c3, c4=c4, c5=c5, c6=c6, c7=c7, c8=c8, c9=c9,
					Pb_PROMOTOR=str(P_Clave), Pb_FEINI=P_Feini, Pb_FEFIN=P_Fefin)
				libro_nombre = P_Clave + "_" + P_Feini.replace("/", "") + "_" + P_Fefin.replace("/",
																								"") + '_comisiones.xlsx'
				for row in c_head:
					has_agent = True
					for i in range(0, len(row) - 4):
						ws.cell(row=4 + i, column=9).value = row[i]
					for i in range(len(row) - 4, len(row)):
						ws.cell(row=i, column=4).value = row[i]
				if not has_agent:
					statement.close()
					c_head.close()
					c1.close()
					c2.close()
					c3.close()
					c4.close()
					c5.close()
					c6.close()
					c7.close()
					c8.close()
					c9.close()
					return make_response(jsonify(succes=False, message="Codigo de promotor no encontrado"), 400)
				c_head.close()
				f = 13 #principal gestor de filas del archivo
				greyFill = PatternFill(fill_type='solid', start_color='d9d9d9', end_color='d9d9d9')
				# NUEVO BLOQUE SECUENCIAL
				cursores = []
				cursores.append(c1)
				cursores.append(c2)
				cursores.append(c3)
				cursores.append(c4)
				cursores.append(c5)
				cursores.append(c6)
				cursores.append(c7)
				cursores.append(c8)
				cursores.append(c9)
				c_count = 1
				for cursor in cursores:
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
					for row in cursor:
						for i in range(0, len(row)):
							valor = row[i]
							ws.cell(row=f, column=i + 1).value = valor
							ws.cell(row=f, column=i + 1).alignment = Alignment(horizontal="center", vertical="center")
							if len(str(valor)) > 17:
								ws.cell(row=f, column=i + 1).font = Font(name='Arial', size=8)
								if len(str(valor)) > 25:
									ws.cell(row=f, column=i + 1).font = Font(name='Arial', size=7)
						f += 1
					cursor.close()
					c_count += 1
				# fin de bloque
				statement.close()
			# fin de bloque
			except Exception as ex:
				app.logger.error(ex)
				statement.close()
				c_head.close()
				c1.close()
				c2.close()
				c3.close()
				c4.close()
				c5.close()
				c6.close()
				c7.close()
				c8.close()
				c9.close()
		except Exception as ex:
			app.logger.error(ex)
			statement.close()
			c_head.close()
			c1.close()
			c2.close()
			c3.close()
			c4.close()
			c5.close()
			c6.close()
			c7.close()
			c8.close()
			c9.close()
		print("\nTermina Proceso " + time.strftime("%X"))

		virtual_wb = BytesIO()
		wb.save(virtual_wb)
		return Response(virtual_wb.getvalue(), mimetype=wb.mime_type,headers={"Content-Disposition": "attachment;filename="+libro_nombre})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="Error en la generacion del archivo"),400)


def getHeadColumns(extension):
    lista = []
    lista.append('# Bono')
    lista.append('Tipo Bono')
    lista.append('Subtipo Bono')
    if extension == "excel":
        lista.append('Grupo')
    lista.append('Oficina')
    lista.append('Ramo')
    lista.append('Poliza')
    lista.append('Contratante')
    if extension == "excel":
        lista.append('Agentes')
        lista.append('Computabilidad')
        lista.append('Causa de No Computabilidad')
    lista.append('Tipo Cambio')
    lista.append('# Recibo')
    lista.append('Serie')
    lista.append('Prima Total')
    lista.append('Prima Neta')
    lista.append('% Bono Pagado')
    lista.append('Monto comisión neta')
    lista.append('Total comisión pagada')
    lista.append('# Liquidación')
    if extension == "excel":
        lista.append('# Comprobante')
    lista.append('Fecha Movimiento')
    return lista


def getHeadColumnsComisones(extension,cursor):
	lista = []
	if cursor in [1,2]:
		lista.append("TIPO")
		lista.append("BASE")
		lista.append("IVA")
		lista.append("SUBTOTAL")
		lista.append("IVARETENIDO")
		lista.append("ISR")
		lista.append("IMP. CEDULAR")
		lista.append("TOTAL")
	if cursor in [3,4,6,7]:
		lista.append("CONCEPTO")
		lista.append("IMPORTE")

	if cursor in [5]:
		lista.append("Daños/Vida")
		if extension =="excel":
			lista.append("Grupo")
		lista.append("Oficina")
		lista.append("Ramo")
		lista.append("Poliza")
		lista.append("Contratante")
		if extension == "excel":
			lista.append("Clave Agente")
		lista.append("Tipo de Cambio")
		lista.append("# Recibo")
		lista.append("Serie de Recibo")
		lista.append("Prima Total")
		lista.append("Prima Neta")
		lista.append("% Comisión pagado")
		lista.append("% Comisión de derecho")
		lista.append("Monto Comisión Neta")
		lista.append("Total Comisión pagado")
		lista.append("# Liquidación")
		if extension == "excel":
			lista.append("# Comprobante")
		lista.append("Fecha aplicación de la póliza")
	if cursor in [8,9]:
		lista.append("FECHA DE LIQUIDACIÓN")
		lista.append("DAÑOS")
		lista.append("VIDA")
		lista.append("TOTAL")
		lista.append("FECHA DE PAGO")
		lista.append("NUMERO DE COMPROBANTE")
		lista.append("IMPORTEPAGADO DAÑOS")
		lista.append("IMPORTE PAGADO VIDA")
		lista.append("TOTAL")
	return lista


def getTipoSubBono(id):
	codigo=""
	if id == 20:
		codigo = "BS"
	if id == 30:
		codigo = "BC"
	if id == 40:
		codigo = "BP"
	if id == 50:
		codigo = "BSE"
	if id == 60:
		codigo = "BCS"
	if id == 70:
		codigo = "BPS"
	if id == 80:
		codigo = "BPF"
	if id == 82:
		codigo = "BD"
	return codigo


def getTableNamesComisiones(tabla):
	nombre=""
	if tabla == 1:
		nombre = "TOTAL DE PERCEPCIONES MENSUALES"
	if tabla == 2:
		nombre = "TOTAL DE PERCEPCIONES ACUMULADO ANUAL"
	if tabla == 3:
		nombre = "DAÑOS MONEDA MXP"
	if tabla == 4:
		nombre = "DAÑOS MONEDA USD"
	if tabla == 5:
		nombre = "DETALLES"
	if tabla == 6:
		nombre = "VIDA MONEDA MXP"
	if tabla == 7:
		nombre = "VIDA MONEDA USD"
	if tabla == 8:
		nombre = "RESUMEN DE DEPOSITOS EN MXP"
	if tabla == 9:
		nombre = "RESUMEN DE DEPOSITOS EN USD"
	return nombre


if __name__ == '__main__':
    app.run(debug=True, port=SERVER_PORT)

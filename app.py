# -*- coding: utf-8 -*-
import os
import cx_Oracle
from flask import Flask, jsonify, request, Response, make_response
import datetime
from dotenv import load_dotenv
from flask_cors import CORS
from logica import getperiodo
from logica import comisiones_xlsx
from logica import comisiones_pdf
from logica import bonos_pdf
from logica import bonos_xlx


app = Flask(__name__)
CORS(app)
app.config['JSON_SORT_KEYS'] = False
context_path = "/api/estados-cuenta/"
load_dotenv()
SERVER_PORT = os.getenv('SERVER_PORT')
# Testing Route
@app.route(context_path+'/ping', methods=['GET'])
def ping():

    return jsonify({'response': 'pong!'})


@app.route(context_path+'/periodo', methods=['GET'])
def periodo():
	P_MES = request.args['mes']
	P_ANIO = request.args['anio']
	P_CLAVE = request.args['clave']
	try:
		estado, mensaje, desde, hasta = getperiodo(P_CLAVE, P_MES, P_ANIO,app)
		if not estado:
			return make_response(jsonify(succes=False, message=mensaje), 400)
		return make_response(jsonify(succes=True, desde=desde, hasta=hasta), 200)
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="La aplicacion tuvo un fallo inesperado."), 400)


@app.route(context_path + '/agentes/bonos/excel', methods=['GET'])
def bono_agente_xlsx():
	P_Clave = request.args['codigo']
	P_Feini = request.args['desde']
	P_Fefin = request.args['hasta']
	P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
	P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
	try:
		estado,mensaje,file, filemime, filename = bonos_xlx(P_Clave, P_Feini, P_Fefin, 'A',app)

		if not estado:
			return make_response(jsonify(succes=False, message=mensaje), 400)
		return Response(file, mimetype=filemime,
						headers={"Content-Disposition": "attachment;filename=" + filename})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="La aplicacion tuvo un fallo inesperado."), 400)


@app.route(context_path + '/agentes/bonos/pdf', methods=['GET'])
def bono_agente_pdf():
	P_Clave = request.args['codigo']
	P_Feini = request.args['desde']
	P_Fefin = request.args['hasta']
	P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
	P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
	try:
		estado,mensaje,file, filemime, filename = bonos_pdf(P_Clave, P_Feini, P_Fefin, 'A',app)

		if not estado:
			return make_response(jsonify(succes=False, message=mensaje), 400)
		return Response(file, mimetype=filemime,
						headers={"Content-Disposition": "attachment;filename=" + filename})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="La aplicacion tuvo un fallo inesperado."), 400)


@app.route(context_path + '/agentes/comisiones/pdf', methods=['GET'])
def comisiones_agente_pdf():
	P_Clave = request.args['codigo']
	P_Feini = request.args['desde']
	P_Fefin = request.args['hasta']
	P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
	P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
	try:
		estado,mensaje,file, filemime, filename = comisiones_pdf(P_Clave, P_Feini, P_Fefin, 'A',app)

		if not estado:
			return make_response(jsonify(succes=False, message=mensaje), 400)
		return Response(file, mimetype=filemime,
						headers={"Content-Disposition": "attachment;filename=" + filename})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="La aplicacion tuvo un fallo inesperado."), 400)


@app.route(context_path + '/agentes/comisiones/excel', methods=['GET'])
def comisiones_agente_xlsx():
	P_Clave = request.args['codigo']
	P_Feini = request.args['desde']
	P_Fefin = request.args['hasta']
	P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
	P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
	try:
		estado,mensaje,file, filemime, filename = comisiones_xlsx(P_Clave, P_Feini, P_Fefin, 'A',app)

		if not estado:
			return make_response(jsonify(succes=False, message=mensaje), 400)
		return Response(file, mimetype=filemime,
						headers={"Content-Disposition": "attachment;filename=" + filename})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="La aplicacion tuvo un fallo inesperado."), 400)


@app.route(context_path + '/promotores/bonos/excel', methods=['GET'])
def bono_promotores_xlsx():
	P_Clave = request.args['codigo']
	P_Feini = request.args['desde']
	P_Fefin = request.args['hasta']
	P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
	P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
	try:
		estado,mensaje,file, filemime, filename = bonos_xlx(P_Clave, P_Feini, P_Fefin, 'P',app)

		if not estado:
			return make_response(jsonify(succes=False, message=mensaje), 400)
		return Response(file, mimetype=filemime,
						headers={"Content-Disposition": "attachment;filename=" + filename})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="La aplicacion tuvo un fallo inesperado."), 400)

@app.route(context_path + '/promotores/bonos/pdf', methods=['GET'])
def bono_promotores_pdf():
	P_Clave = request.args['codigo']
	P_Feini = request.args['desde']
	P_Fefin = request.args['hasta']
	P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
	P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
	try:
		estado,mensaje,file, filemime, filename = bonos_pdf(P_Clave, P_Feini, P_Fefin, 'P',app)

		if not estado:
			return make_response(jsonify(succes=False, message=mensaje), 400)
		return Response(file, mimetype=filemime,
						headers={"Content-Disposition": "attachment;filename=" + filename})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="La aplicacion tuvo un fallo inesperado."), 400)


@app.route(context_path + '/promotores/comisiones/excel', methods=['GET'])
def comisiones_promotor_xlsx():
	P_Clave = request.args['codigo']
	P_Feini = request.args['desde']
	P_Fefin = request.args['hasta']
	P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
	P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
	try:
		estado,mensaje,file, filemime, filename = comisiones_xlsx(P_Clave, P_Feini, P_Fefin, 'P',app)

		if not estado:
			return make_response(jsonify(succes=False, message=mensaje), 400)
		return Response(file, mimetype=filemime,
						headers={"Content-Disposition": "attachment;filename=" + filename})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="La aplicacion tuvo un fallo inesperado."), 400)


@app.route(context_path + '/promotores/comisiones/pdf', methods=['GET'])
def comisiones_promotores_pdf():
	P_Clave = request.args['codigo']
	P_Feini = request.args['desde']
	P_Fefin = request.args['hasta']
	P_Feini = datetime.datetime.strptime(P_Feini, "%Y-%m-%d").strftime("%d/%m/%Y")
	P_Fefin = datetime.datetime.strptime(P_Fefin, "%Y-%m-%d").strftime("%d/%m/%Y")
	try:
		estado,mensaje,file, filemime, filename=comisiones_pdf(P_Clave,P_Feini,P_Fefin,'P',app)

		if not estado:
			return make_response(jsonify(succes=False, message=mensaje), 400)
		return Response(file, mimetype=filemime,
						headers={"Content-Disposition": "attachment;filename=" + filename})
	except Exception as ex:
		app.logger.error(ex)
		return make_response(jsonify(succes=False, message="La aplicacion tuvo un fallo inesperado."), 400)


if __name__ == '__main__':
    app.run(debug=True, port=SERVER_PORT)

# -*- coding: utf-8 -*-
from openerp import models, fields, api, _
import time
from itertools import ifilter

class account_voucher(models.Model):
	_inherit	= "account.voucher"

	cai_shot	= fields.Char("Cai", readonly=True)
	cai_expires_shot= fields.Date("expiration_date", readonly=True)
	min_number_shot = fields.Integer("min_number", readonly=True)
	max_number_shot = fields.Integer("max_number", readonly=True)

	amount_total_text = fields.Char("Amount Total", compute = 'get_totalt', default='Cero')


	def proforma_voucher(self, cr, uid, ids, context=None):
		""" La fecha de la factura debe estar en el rango, si se esta usando """
		voucher = self.pool.get('account.voucher').browse(cr, uid, ids)
		if voucher.journal_id.sequence_id.fiscal_regime:
			if voucher.date > voucher.journal_id.sequence_id.expiration_date:
				voucher.journal_id.sequence_id.number_next_actual = voucher.journal_id.sequence_id.number_next_actual -1
				raise Warning(_('la fecha de expiración para esta secuencia es %s ') %(voucher.journal_id.sequence_id.expiration_date) )
			voucher.cai_shot=''

			for regimen in voucher.journal_id.sequence_id.fiscal_regime:
				if regimen.selected:
					voucher.cai_shot = regimen.cai.name
					voucher.cai_expires_shot = regimen.cai.expiration_date
					voucher.min_number_shot = regimen.desde
					voucher.max_number_shot = regimen.hasta	
        	self.action_move_line_create(cr, uid, ids, context=context)
        	return True


	@api.one
	@api.depends('journal_id')
	def get_totalt(self):
		self.amount_total_text=''
		if self.currency_id:
			self.amount_total_text=self.to_word(self.amount,self.currency_id.name)
		else:
			self.amount_total_text =self.to_word(self.amount,self.user_id.company_id.currency_id.name)
		return True

	def to_word(self,number, mi_moneda):
		valor= number
		number=int(number)
		centavos=int((round(valor-number,2))*100)
		UNIDADES = (
			'',
			'UN ',
			'DOS ',
			'TRES ',
			'CUATRO ',
			'CINCO ',
			'SEIS ',
			'SIETE ',
			'OCHO ',
			'NUEVE ',
			'DIEZ ',
			'ONCE ',
			'DOCE ',
			'TRECE ',
			'CATORCE ',
			'QUINCE ',
			'DIECISEIS ',
			'DIECISIETE ',
			'DIECIOCHO ',
			'DIECINUEVE ',
			'VEINTE '
		)

		DECENAS = (
			'VENTI',
			'TREINTA ',
			'CUARENTA ',
			'CINCUENTA ',
			'SESENTA ',
			'SETENTA ',
			'OCHENTA ',
			'NOVENTA ',
			'CIEN ')

		CENTENAS = (
			'CIENTO ',
			'DOSCIENTOS ',
			'TRESCIENTOS ',
			'CUATROCIENTOS ',
			'QUINIENTOS ',
			'SEISCIENTOS ',
			'SETECIENTOS ',
			'OCHOCIENTOS ',
			'NOVECIENTOS '
		)
		MONEDAS = (
			{'country': u'Colombia', 'currency': 'COP', 'singular': u'PESO COLOMBIANO', 'plural': u'PESOS COLOMBIANOS', 'symbol': u'$'},
			{'country': u'Honduras', 'currency': 'HNL', 'singular': u'Lempira', 'plural': u'Lempiras', 'symbol': u'L'},
			{'country': u'Estados Unidos', 'currency': 'USD', 'singular': u'DÓLAR', 'plural': u'DÓLARES', 'symbol': u'US$'},
			{'country': u'Europa', 'currency': 'EUR', 'singular': u'EURO', 'plural': u'EUROS', 'symbol': u'€'},
			{'country': u'México', 'currency': 'MXN', 'singular': u'PESO MEXICANO', 'plural': u'PESOS MEXICANOS', 'symbol': u'$'},
			{'country': u'Perú', 'currency': 'PEN', 'singular': u'NUEVO SOL', 'plural': u'NUEVOS SOLES', 'symbol': u'S/.'},
			{'country': u'Reino Unido', 'currency': 'GBP', 'singular': u'LIBRA', 'plural': u'LIBRAS', 'symbol': u'£'}
			)
		if mi_moneda != None:
			try:
				moneda = ifilter(lambda x: x['currency'] == mi_moneda, MONEDAS).next()
				if number < 2:
					moneda = moneda['singular']
				else:
					moneda = moneda['plural']
			except:
				return "Tipo de moneda inválida"
		else:
			moneda = ""
		converted = ''
		if not (0 < number < 999999999):
			return 'No es posible convertir el numero a letras'

		number_str = str(number).zfill(9)
		millones = number_str[:3]
		miles = number_str[3:6]
		cientos = number_str[6:]

		if(millones):
			if(millones == '001'):
				converted += 'UN MILLON '
			elif(int(millones) > 0):
				converted += '%sMILLONES ' % self.convert_group(millones)

		if(miles):
			if(miles == '001'):
				converted += 'MIL '
			elif(int(miles) > 0):
				converted += '%sMIL ' % self.convert_group(miles)

		if(cientos):
			if(cientos == '001'):
				converted += 'UN '
			elif(int(cientos) > 0):
				converted += '%s ' % self.convert_group(cientos)
		if(centavos)>0:
			converted+= "con %2i/100 "%centavos
		converted += moneda
		return converted.title()

	def convert_group(self,n):
		UNIDADES = (
			'',
			'UN ',
			'DOS ',
			'TRES ',
			'CUATRO ',
			'CINCO ',
			'SEIS ',
			'SIETE ',
			'OCHO ',
			'NUEVE ',
			'DIEZ ',
			'ONCE ',
			'DOCE ',
			'TRECE ',
			'CATORCE ',
			'QUINCE ',
			'DIECISEIS ',
			'DIECISIETE ',
			'DIECIOCHO ',
			'DIECINUEVE ',
			'VEINTE '
		)
		DECENAS = (
			'VENTI',
			'TREINTA ',
			'CUARENTA ',
			'CINCUENTA ',
			'SESENTA ',
			'SETENTA ',
			'OCHENTA ',
			'NOVENTA ',
			'CIEN '
		)

		CENTENAS = (
			'CIENTO ',
			'DOSCIENTOS ',
			'TRESCIENTOS ',
			'CUATROCIENTOS ',
			'QUINIENTOS ',
			'SEISCIENTOS ',
			'SETECIENTOS ',
			'OCHOCIENTOS ',
			'NOVECIENTOS '
		)
		MONEDAS = (
			{'country': u'Colombia', 'currency': 'COP', 'singular': u'PESO COLOMBIANO', 'plural': u'PESOS COLOMBIANOS', 'symbol': u'$'},
			{'country': u'Honduras', 'currency': 'HNL', 'singular': u'Lempira', 'plural': u'Lempiras', 'symbol': u'L'},
			{'country': u'Estados Unidos', 'currency': 'USD', 'singular': u'DÓLAR', 'plural': u'DÓLARES', 'symbol': u'US$'},
			{'country': u'Europa', 'currency': 'EUR', 'singular': u'EURO', 'plural': u'EUROS', 'symbol': u'€'},
			{'country': u'México', 'currency': 'MXN', 'singular': u'PESO MEXICANO', 'plural': u'PESOS MEXICANOS', 'symbol': u'$'},
			{'country': u'Perú', 'currency': 'PEN', 'singular': u'NUEVO SOL', 'plural': u'NUEVOS SOLES', 'symbol': u'S/.'},
			{'country': u'Reino Unido', 'currency': 'GBP', 'singular': u'LIBRA', 'plural': u'LIBRAS', 'symbol': u'£'}
		)
		output = ''

		if(n == '100'):
			output = "CIEN "
		elif(n[0] != '0'):
			output = CENTENAS[int(n[0]) - 1]

		k = int(n[1:])
		if(k <= 20):
			output += UNIDADES[k]
		else:
			if((k > 30) & (n[2] != '0')):
				output += '%sY %s' % (DECENAS[int(n[1]) - 2], UNIDADES[int(n[2])])
			else:
				output += '%s%s' % (DECENAS[int(n[1]) - 2], UNIDADES[int(n[2])])

		return output

	def addComa(self, snum ):
		s = snum;
		i = s.index('.') # Se busca la posición del punto decimal
		while i > 3:
			i = i - 3
			s = s[:i] +  ',' + s[i:]
		return s


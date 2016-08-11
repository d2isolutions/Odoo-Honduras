# -*- coding: utf-8 -*-
from openerp import models, fields, api, _
import time
from openerp.exceptions import except_orm, Warning, RedirectWarning

class fiscal_regimen(models.Model):

	_name			="dei.fiscal_regime"
	cai 			= fields.Many2one('dei.cai', required=True)
	sequence		= fields.Many2one('ir.sequence')
	selected		= fields.Boolean('selected')
	desde			= fields.Integer('From')
	hasta			= fields.Integer('to')


	@api.onchange('selected')
	def disable_other_regimes(self):
		if self.selected:
			lista = self.env['dei.fiscal_regime'].search([('sequence.name','=',self.sequence.name)])
			for regime in lista:
				regime.write({'selected':0})
			self.write({'selected':1})



class cai(models.Model):
	_name 			= "dei.cai"
	name 			= fields.Char('CAI', help='Clave de Autorización de Impresion ', required=True, select=True)
	expiration_date 	= fields.Date('Expiration Date', required=True, select=True)
	company			= fields.Many2one('res.company', required=True)
	fiscal_regimes		= fields.One2many('dei.fiscal_regime','cai')


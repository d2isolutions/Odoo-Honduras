from openerp import models, fields, _, api

class res_partner(models.Model):
	_inherit = 'res.partner'
	
	rtn =fields.Char('RTN',translate=True)

	_default = {
        'lang': 'es_HN',
	}

	_sql_constraints = [('rtn_uniq','unique(rtn)', 'rtn1 must be unique!')]




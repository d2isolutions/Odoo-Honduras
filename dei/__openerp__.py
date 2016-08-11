# -*- coding: utf-8 -*-
{
	'name'		:'DEI HONDURAS',
	'version'	:'2.0',
	'author'	:'Alejandro Rodriguez',
	'description'	:"""
					Honduras new fiscal regimen 
	""",
	'depends'	:['base','account'],
	'data'		:[
			'views/cai_view.xml',
			'views/ir_sequence_view.xml',
			'views/res_partner_view.xml',
			'reports/report_sales_receipt.xml',
			],

	'update_xml'	:[
			'security/groups.xml',
                       'security/ir.model.access.csv',
			],
	'installable':True,
}


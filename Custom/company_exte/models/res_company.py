from odoo import models, fields, api

FREE_ZONE_IMPLEMENTING_REGULATIONS = {
    "Abu Dhabi Global Market": "Abu Dhabi Global Market Companies Regulations 2020",
    "Ajman Free Zone": "Ajman Free Zone Companies Regulations (Decree No. 9 of 2016)",
    "Creative Media Authority": "Emirate of Abu Dhabi Law No 8 of 2022 relating to the Creative Zone (Law 8)",
    "Department of Economic Development": "Federal Decree Law No. (32) of 2021 on Commercial Companies",
    "Department of Economy and Tourism": "Federal Decree Law No. (32) of 2021 on Commercial Companies",
    "Dubai Aviation City Corporation": "Law No. (10) of 2015 concerning the Dubai Aviation City Corporation",
    "Dubai Development Authority Zone": "Law No. (10) of 2018 on the Dubai Development Authority's Organization Structure",
    "Dubai International Financial Centre": "Dubai International Financial Centre Companies Law No. 5 of 2018",
    "Dubai Integrated Economic Zones Authority": "Dubai Integrated Economic Zones Authority Implementing Regulations 2023",
    "Dubai Multi Commodities Centre Free Zone": "Dubai Multi Commodities Centre Authority Company Regulations 2024",
    "Dubai World Trade Centre Authority": "Dubai World Trade Centre Authority Company Regulations 2015",
    "Meydan Free Zone": "Meydan Free Zone Companies and licensing Regulations 2022",
    "Ports, Customs and Free Zone Corporation": "Law No. (02) of 2023 concerning the Dubai Ports Authority",
    "Creative City Media Free Zone": "Company Regulations and Licensing Regulations of the City of Fujairah for Creative Zone - Free Zones 2024",
    "Ras Al Khaimah Economic Zone": "Ras Al Khaimah Economic Zone Companies Regulations 2023",
    "Ras Al Khaimah International Corporate Centre": "Ras Al Khaimah International Corporate Centre Business Companies Regulations 2018",
    "Sharjah Media City": "Sharjah Media City Free Zone Authority Companies and Licensing Regulations 2024",
    "Sharjah Publishing City": "Sharjah Publishing City, Free Zone Authority Implementing Regulations 2023",
    "Sharjah Research Technology and Innovation Park Free Zone Authority": (
        "Emiri Decree No. (09) of 2023, Sharjah Research Technology and Innovation Park Free Zone Authority"
    ),
}

FREE_ZONE_LOCATION_DEFAULTS = {
    "Dubai Integrated Economic Zones Authority": {
        'street': "DSO-IFZA, IFZA Properties, Dubai Silicon Oasis",
        'city': "Dubai",
    },
    "Meydan Free Zone": {
        'street': "Meydan Grandstand, 6th Floor, Meydan Road, Nad Al Sheba",
        'city': "Dubai",
    },
    "Dubai International Financial Centre": {
        'city': "Dubai",
    },
    "Sharjah Media City": {
        'street': "Sharjah Media City",
        'city': "Sharjah",
    },
    "Creative City Media Free Zone": {
        'street': "Fujairah - Creative Tower, P.O.Box 4422",
        'city': "Fujairah",
    },
    "Sharjah Publishing City": {
        'street': "Business Centre, Sharjah Publishing City Freezone",
        'city': "Sharjah",
    },
}

FREE_ZONE_ALIASES = {
    "Ajman Free Zone, Ajman": "Ajman Free Zone",
    "Ajman Free zone": "Ajman Free Zone",
    "Dubai Integrated Economic Zones": "Dubai Integrated Economic Zones Authority",
    "Creative Media City": "Creative City Media Free Zone",
}

class ResCompany(models.Model):
    _inherit = 'res.company'

    # Free Zone & License
    free_zone = fields.Selection(
        [
            ("Abu Dhabi Global Market", "Abu Dhabi Global Market"),
            ("Ajman Free Zone", "Ajman Free Zone, Ajman"),
            ("Creative Media Authority", "Creative Media Authority"),
            ("Department of Economic Development", "Department of Economic Development"),
            ("Department of Economy and Tourism", "Department of Economy and Tourism"),
            ("Dubai Aviation City Corporation", "Dubai Aviation City Corporation"),
            ("Dubai Development Authority Zone", "Dubai Development Authority Zone"),
            ("Dubai International Financial Centre", "Dubai International Financial Centre"),
            ("Dubai Integrated Economic Zones Authority", "Dubai Integrated Economic Zones Authority"),
            ("Dubai Multi Commodities Centre Free Zone", "Dubai Multi Commodities Centre Free Zone"),
            ("Dubai World Trade Centre Authority", "Dubai World Trade Centre Authority"),
            ("Meydan Free Zone", "Meydan Free Zone"),
            ("Ports, Customs and Free Zone Corporation", "Ports, Customs and Free Zone Corporation"),
            ("Creative City Media Free Zone", "Creative City Media Free Zone"),
            ("Ras Al Khaimah Economic Zone", "Ras Al Khaimah Economic Zone"),
            ("Ras Al Khaimah International Corporate Centre", "Ras Al Khaimah International Corporate Centre"),
            ("Sharjah Media City", "Sharjah Media City"),
            ("Sharjah Publishing City", "Sharjah Publishing City"),
            ("Sharjah Research Technology and Innovation Park Free Zone Authority", "Sharjah Research Technology and Innovation Park Free Zone Authority"),
        ],
        string="Company Free Zone",
    )
    company_license_number = fields.Char(string="Company License Number")
    trade_license_activities = fields.Text(string="Trade License Activities")

    # Tax Information
    corporate_tax_registration_number = fields.Char(string="Corporate Tax Registration Number")
    vat_registration_number = fields.Char(string="VAT Registration Number")
    corporate_tax_start_date = fields.Date(string="Corporate Tax Start Date")
    corporate_tax_end_date = fields.Date(string="Corporate Tax End Date")

    # Incorporation
    incorporation_date = fields.Date(string="Company Incorporation Date")

    # Regulations
    implementing_regulations_freezone = fields.Text(string="Implementing Regulations for Free Zones")

    company_share = fields.Selection(
        [('paid', 'Paid'), ('unpaid', 'Unpaid')],
        string='Share',
    )

    @api.model
    def _get_free_zone_implementing_regulations(self, free_zone):
        normalized = self._normalize_free_zone(free_zone)
        return FREE_ZONE_IMPLEMENTING_REGULATIONS.get(normalized or '')

    @api.model
    def _get_free_zone_location_defaults(self, free_zone):
        normalized = self._normalize_free_zone(free_zone)
        defaults = dict(FREE_ZONE_LOCATION_DEFAULTS.get(normalized or '', {}))
        if defaults.get('city'):
            return defaults
        if normalized == 'Abu Dhabi Global Market':
            defaults['city'] = 'Abu Dhabi'
        elif normalized == 'Ajman Free Zone':
            defaults['city'] = 'Ajman'
        elif normalized in ('Department of Economic Development', 'Department of Economy and Tourism'):
            defaults['city'] = 'Dubai'
        elif 'Dubai' in (normalized or ''):
            defaults['city'] = 'Dubai'
        elif 'Ras Al Khaimah' in (normalized or ''):
            defaults['city'] = 'Ras Al Khaimah'
        elif 'Sharjah' in (normalized or ''):
            defaults['city'] = 'Sharjah'
        return defaults

    @api.model
    def _normalize_free_zone(self, free_zone):
        if not free_zone:
            return free_zone
        return FREE_ZONE_ALIASES.get(free_zone, free_zone)

    @api.model_create_multi
    def create(self, vals_list):
        for vals in vals_list:
            if 'free_zone' in vals:
                vals['free_zone'] = self._normalize_free_zone(vals.get('free_zone'))
        return super().create(vals_list)

    def write(self, vals):
        if 'free_zone' in vals:
            vals = dict(vals)
            vals['free_zone'] = self._normalize_free_zone(vals.get('free_zone'))
        return super().write(vals)

    @api.onchange('free_zone')
    def _onchange_free_zone(self):
        if not self.free_zone:
            self.implementing_regulations_freezone = False
            return
        mapped = self._get_free_zone_implementing_regulations(self.free_zone)
        if mapped:
            self.implementing_regulations_freezone = mapped
        location_defaults = self._get_free_zone_location_defaults(self.free_zone)
        if location_defaults.get('street'):
            self.street = location_defaults['street']
        if location_defaults.get('city'):
            self.city = location_defaults['city']


    # Individual Shareholder Fields (1 to 10)
    shareholder_1 = fields.Char(string="Shareholder 1")
    shareholder_2 = fields.Char(string="Shareholder 2")
    shareholder_3 = fields.Char(string="Shareholder 3")
    shareholder_4 = fields.Char(string="Shareholder 4")
    shareholder_5 = fields.Char(string="Shareholder 5")
    shareholder_6 = fields.Char(string="Shareholder 6")
    shareholder_7 = fields.Char(string="Shareholder 7")
    shareholder_8 = fields.Char(string="Shareholder 8")
    shareholder_9 = fields.Char(string="Shareholder 9")
    shareholder_10 = fields.Char(string="Shareholder 10")

    nationality_1 = fields.Char(string="Nationality 1")
    nationality_2 = fields.Char(string="Nationality 2")
    nationality_3 = fields.Char(string="Nationality 3")
    nationality_4 = fields.Char(string="Nationality 4")
    nationality_5 = fields.Char(string="Nationality 5")
    nationality_6 = fields.Char(string="Nationality 6")
    nationality_7 = fields.Char(string="Nationality 7")
    nationality_8 = fields.Char(string="Nationality 8")
    nationality_9 = fields.Char(string="Nationality 9")
    nationality_10 = fields.Char(string="Nationality 10")

    # No. of Shares per Shareholder
    number_of_shares_1 = fields.Integer(string="No. of Shares 1")
    number_of_shares_2 = fields.Integer(string="No. of Shares 2")
    number_of_shares_3 = fields.Integer(string="No. of Shares 3")
    number_of_shares_4 = fields.Integer(string="No. of Shares 4")
    number_of_shares_5 = fields.Integer(string="No. of Shares 5")
    number_of_shares_6 = fields.Integer(string="No. of Shares 6")
    number_of_shares_7 = fields.Integer(string="No. of Shares 7")
    number_of_shares_8 = fields.Integer(string="No. of Shares 8")
    number_of_shares_9 = fields.Integer(string="No. of Shares 9")
    number_of_shares_10 = fields.Integer(string="No. of Shares 10")

    # Share Value per Shareholder
    share_value_1 = fields.Float(string="Share Value 1")
    share_value_2 = fields.Float(string="Share Value 2")
    share_value_3 = fields.Float(string="Share Value 3")
    share_value_4 = fields.Float(string="Share Value 4")
    share_value_5 = fields.Float(string="Share Value 5")
    share_value_6 = fields.Float(string="Share Value 6")
    share_value_7 = fields.Float(string="Share Value 7")
    share_value_8 = fields.Float(string="Share Value 8")
    share_value_9 = fields.Float(string="Share Value 9")
    share_value_10 = fields.Float(string="Share Value 10")

    # Computed Total per Shareholder
    total_share_1 = fields.Float(string="Total Share 1", compute='_compute_total_shares', store=True, readonly=True)
    total_share_2 = fields.Float(string="Total Share 2", compute='_compute_total_shares', store=True, readonly=True)
    total_share_3 = fields.Float(string="Total Share 3", compute='_compute_total_shares', store=True, readonly=True)
    total_share_4 = fields.Float(string="Total Share 4", compute='_compute_total_shares', store=True, readonly=True)
    total_share_5 = fields.Float(string="Total Share 5", compute='_compute_total_shares', store=True, readonly=True)
    total_share_6 = fields.Float(string="Total Share 6", compute='_compute_total_shares', store=True, readonly=True)
    total_share_7 = fields.Float(string="Total Share 7", compute='_compute_total_shares', store=True, readonly=True)
    total_share_8 = fields.Float(string="Total Share 8", compute='_compute_total_shares', store=True, readonly=True)
    total_share_9 = fields.Float(string="Total Share 9", compute='_compute_total_shares', store=True, readonly=True)
    total_share_10 = fields.Float(string="Total Share 10", compute='_compute_total_shares', store=True, readonly=True)

    # Compute individual total shares
    @api.depends(
        'number_of_shares_1', 'share_value_1',
        'number_of_shares_2', 'share_value_2',
        'number_of_shares_3', 'share_value_3',
        'number_of_shares_4', 'share_value_4',
        'number_of_shares_5', 'share_value_5',
        'number_of_shares_6', 'share_value_6',
        'number_of_shares_7', 'share_value_7',
        'number_of_shares_8', 'share_value_8',
        'number_of_shares_9', 'share_value_9',
        'number_of_shares_10', 'share_value_10',
    )
    def _compute_total_shares(self):
        for rec in self:
            for i in range(1, 11):
                shares = rec[f'number_of_shares_{i}'] or 0
                value = rec[f'share_value_{i}'] or 0.0
                rec[f'total_share_{i}'] = shares * value

    # Compute GRAND TOTAL SHARE
    @api.depends(
        'total_share_1', 'total_share_2', 'total_share_3', 'total_share_4',
        'total_share_5', 'total_share_6', 'total_share_7', 'total_share_8',
        'total_share_9', 'total_share_10'
    )
    def _compute_total_share(self):
        for rec in self:
            rec.total_share = sum(rec[f'total_share_{i}'] for i in range(1, 11))
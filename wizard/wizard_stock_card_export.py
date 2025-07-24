from odoo import models, fields
from datetime import date
import base64
from io import BytesIO
from xlsxwriter.workbook import Workbook
from psycopg2 import sql

class WizardStockCardExport(models.TransientModel):
    _name = 'wizard.stock.card.export'
    _description = 'Xuất Thẻ Kho'

    product_id = fields.Many2one('product.product', string='Vật tư', required=True)
    from_date = fields.Date(string="Từ ngày", required=True)
    to_date = fields.Date(string="Đến ngày", required=True)
    company_id = fields.Many2one('res.company', string='Công ty', default=lambda self: self.env.company, required=True)
    file = fields.Binary()
    file_name = fields.Char()

    def action_export_stock_card(self):
        self.ensure_one()
        report_xlsx_data = self.generate_report_xlsx()
        self.write({
            'file_name': 'Thẻ Kho.xlsx',
            'file': report_xlsx_data
        })
        return {
            'name': 'Thẻ Kho',
            'type': 'ir.actions.act_url',
            'url': "web/content/?model=wizard.stock.card.export&id=" + str(self.id) + "&filename_field=file_name&field=file&download=true&filename=" + self.file_name,
        }

    def generate_report_xlsx(self):
        buf = BytesIO()
        wb = Workbook(buf)
        wb.formats[0].font_name = 'Times New Roman'
        wb.formats[0].font_size = 11
        self._generate_xlsx_report(wb)
        wb.close()
        buf.seek(0)
        xlsx_data = buf.getvalue()
        file_data = base64.encodebytes(xlsx_data)
        return file_data

    def _get_ton_dau_ky(self, product_id, company_id, from_date):
        query = """
            SELECT
                COALESCE(SUM(
                    CASE
                        WHEN dest.usage = 'internal' THEN sml.quantity
                        WHEN source.usage = 'internal' THEN -sml.quantity
                        ELSE 0
                    END
                ), 0.0) AS ton_dau_ky
            FROM stock_move_line sml
            LEFT JOIN stock_location source ON sml.location_id = source.id
            LEFT JOIN stock_location dest ON sml.location_dest_id = dest.id
            JOIN stock_move sm ON sml.move_id = sm.id
            WHERE sml.product_id = %s
              AND sml.company_id = %s
              AND sml.date < %s
              AND sm.state = 'done'
        """
        self._cr.execute(query, (product_id, company_id, from_date))
        result = self._cr.fetchone()
        return result[0] or 0.0

    def _get_phat_sinh_trong_ky(self, product_id, company_id, from_date, to_date):
        query = """
            SELECT
                COALESCE(SUM(CASE WHEN dest.usage = 'internal' THEN sml.quantity ELSE 0 END), 0.0) AS qty_in,
                COALESCE(SUM(CASE WHEN source.usage = 'internal' THEN sml.quantity ELSE 0 END), 0.0) AS qty_out
            FROM stock_move_line sml
            LEFT JOIN stock_location source ON sml.location_id = source.id
            LEFT JOIN stock_location dest ON sml.location_dest_id = dest.id
            LEFT JOIN stock_move sm ON sml.move_id = sm.id
            WHERE sml.product_id = %s
              AND sml.company_id = %s
              AND sml.date >= %s AND sml.date <= %s
              AND sm.state = 'done';
        """
        self._cr.execute(query, (product_id, company_id, from_date, to_date))
        result = self._cr.fetchone()
        qty_in = result[0] or 0.0
        qty_out = result[1] or 0.0
        return qty_in, qty_out

    def _generate_xlsx_report(self, workbook):
        sheet = workbook.add_worksheet('CDKT')
        title_style = workbook.add_format({
            'bold': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 16,
        })
        style_normal = workbook.add_format({
            'text_wrap': True,
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 11,
        })
        table_header = workbook.add_format({
            'bold': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': '#C0C0C0',
            'font_name': 'Times New Roman',
            'font_size': 11,
        })
        style_center = workbook.add_format({
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Times New Roman',
            'font_size': 11,
        })
        style_bold = workbook.add_format({
            'bold': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'border': 1,
            'font_name': 'Times New Roman',
            'font_size': 11,
        })
        style_normal_cell = workbook.add_format({
            'text_wrap': True,
            'valign': 'vcenter',
            'border': 1,
            'font_name': 'Times New Roman',
            'font_size': 11,
        })
        style_center_bold = workbook.add_format({
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'bold': 1,
            'border': 1,
            'font_name': 'Times New Roman',
            'font_size': 11,
        })
        style_center_normal = workbook.add_format({
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'font_name': 'Times New Roman',
            'font_size': 11,
        })
        style_right = workbook.add_format({
            'text_wrap': True,
            'align': 'right',
            'valign': 'vcenter',
            'border': 1,
            'font_name': 'Times New Roman',
            'font_size': 11,
        })
        style_right_bold = workbook.add_format({
            'text_wrap': True,
            'bold': 1,
            'align': 'right',
            'valign': 'vcenter',
            'border': 1,
            'font_name': 'Times New Roman',
            'font_size': 11,
        })

        product_id = self.product_id.id
        company_id = self.company_id.id
        from_date = self.from_date
        to_date = self.to_date

        ton_dau_ky = self._get_ton_dau_ky(product_id, company_id, from_date)
        nhap_trong_ky, xuat_trong_ky = self._get_phat_sinh_trong_ky(product_id, company_id, from_date, to_date)

        if ton_dau_ky >= 0:
            ton_dau_ky_nhap = ton_dau_ky
            ton_dau_ky_xuat = ''
        else:
            ton_dau_ky_nhap = ''
            ton_dau_ky_xuat = abs(ton_dau_ky)

        ton_cuoi_ky = ton_dau_ky + nhap_trong_ky - xuat_trong_ky

        if ton_cuoi_ky >= 0:
            ton_cuoi_ky_nhap = ton_cuoi_ky
            ton_cuoi_ky_xuat = ''
        else:
            ton_cuoi_ky_nhap = ''
            ton_cuoi_ky_xuat = abs(ton_cuoi_ky)

        for row in range(6, 11):
            for col in range(0, 8):
                sheet.write(row, col, '', style_normal_cell)

        sheet.merge_range('A1:D1', f"Công ty: {self.company_id}", style_normal)
        sheet.merge_range('A2:D2', 'Địa chỉ: ', style_normal)

        sheet.merge_range('A4:H4', 'THẺ KHO', title_style)
        sheet.merge_range('A5:H5', f"Vật tư: {self.product_id}", style_center)
        sheet.merge_range('A6:H6',f"Từ ngày: {self.from_date.strftime('%d/%m/%Y')}  Đến ngày: {self.to_date.strftime('%d/%m/%Y')}", style_center)

        sheet.merge_range('A7:B7', 'Chứng từ', style_center_normal)
        sheet.write('A8', 'Ngày', style_center_normal)
        sheet.write('B8', 'Số', style_center_normal)
        sheet.merge_range('C7:C8', 'Công trình/Khách hàng', style_center_bold)
        sheet.merge_range('D7:D8', 'Diễn giải', style_center_bold)
        sheet.merge_range('E7:E8', 'Mã NX', style_center_bold)
        sheet.merge_range('F7:F8', 'SL Nhập', style_center_bold)
        sheet.merge_range('G7:G8', 'SL Xuất', style_center_bold)
        sheet.merge_range('H7:H8', 'Tồn Kho', style_center_bold)

        sheet.write('D9', 'Tồn đầu kỳ:', style_bold)
        sheet.write('D10', 'Phát sinh trong kỳ:', style_bold)
        sheet.write('D11', 'Tồn cuối kỳ:', style_bold)

        sheet.write('F9', ton_dau_ky_nhap, style_right_bold)
        sheet.write('F10', nhap_trong_ky, style_right_bold)
        sheet.write('F11', ton_cuoi_ky_nhap, style_right_bold)

        sheet.write('G9', ton_dau_ky_xuat, style_right_bold)
        sheet.write('G10', xuat_trong_ky, style_right_bold)
        sheet.write('G11', ton_cuoi_ky_xuat, style_right_bold)

        row = 11
        ton_kho = ton_dau_ky

        move_lines = self.env['stock.move.line'].search([
            ('product_id', '=', product_id),
            ('company_id', '=', company_id),
            ('date', '>=', from_date),
            ('date', '<=', to_date),
            ('move_id.state', '=', 'done'),
        ], order='date asc')

        for move in move_lines:
            date_str = move.date.strftime('%d-%b-%Y') if move.date else ''
            so_chung_tu = move.picking_id.name if move.picking_id else ''
            doi_tuong = move.picking_id.partner_id.name if move.picking_id and move.picking_id.partner_id else ''
            dien_giai = move.move_id.name or ''
            ma_nx = move.move_id.account_id.code if hasattr(move.move_id, 'account_id') and move.move_id.account_id else ''

            is_nhap = move.location_dest_id.usage == 'internal' and move.location_id.usage != 'internal'
            is_xuat = move.location_id.usage == 'internal' and move.location_dest_id.usage != 'internal'

            sl_nhap = move.quantity if is_nhap else 0.0
            sl_xuat = move.quantity if is_xuat else 0.0

            ton_kho = ton_kho + sl_nhap - sl_xuat

            # Ghi vào Excel
            sheet.write(row, 0, date_str, style_center_normal)
            sheet.write(row, 1, so_chung_tu, style_center_normal)
            sheet.write(row, 2, doi_tuong, style_normal_cell)
            sheet.write(row, 3, dien_giai, style_normal_cell)
            sheet.write(row, 4, ma_nx, style_center_normal)
            sheet.write_number(row, 5, sl_nhap, style_right)
            sheet.write_number(row, 6, sl_xuat, style_right)
            sheet.write_number(row, 7, ton_kho, style_right)

            row += 1

        sheet.write(row, 4, 'TỔNG CỘNG:', style_right_bold)
        sheet.write_formula(row, 5, f"SUM(F12:F{row})", style_right_bold)
        sheet.write_formula(row, 6, f"SUM(G12:G{row})", style_right_bold)
        sheet.write(row, 7, ton_kho, style_right_bold)





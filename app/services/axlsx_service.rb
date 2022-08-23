class AxlsxService
  def initialize(reports)
    @reports = reports
  end

  def create_report
    Axlsx::Package.new do |p|
      wb = p.workbook
      wb.styles do |style|
        yellow_col = style.add_style(bg_color: 'CCCC66', b: true, border: Axlsx::STYLE_THIN_BORDER,
                                     alignment: { horizontal: :center, vertical: :center, wrap_text: true })
        green_col = style.add_style(bg_color: '99CC99', b: true, border: Axlsx::STYLE_THIN_BORDER,
                                    alignment: { horizontal: :center, vertical: :center, wrap_text: true })
        pink_red_col = style.add_style(fg_color: 'FF0000', b: true, bg_color: 'FF99FF',
                                       border: Axlsx::STYLE_THIN_BORDER, alignment: { horizontal: :center, vertical: :center, wrap_text: true })
        pink_col = style.add_style(bg_color: 'FF99FF', b: true, border: Axlsx::STYLE_THIN_BORDER,
                                   alignment: { horizontal: :center, vertical: :center, wrap_text: true })
        grey_col = style.add_style(bg_color: 'A9A9A9', b: true, border: Axlsx::STYLE_THIN_BORDER,
                                   alignment: { horizontal: :center, vertical: :center, wrap_text: true })
        text_color = style.add_style(fg_color: 'FF0000', alignment: { horizontal: :center })
        main_text = style.add_style(b: true, alignment: { horizontal: :center })
        left_alignment = style.add_style(alignment: { horizontal: :left })
        center_alignment = style.add_style(alignment: { horizontal: :center })
        right_alignment = style.add_style(alignment: { horizontal: :right })
        right_alignment_with_thin_border = style.add_style(alignment: { horizontal: :right },
                                                           border: {
                                                             style: :thin, color: '000000', edges: [:bottom]
                                                           })
        green_col_with_thick_border = style.add_style(bg_color: '99CC99', b: true, border: { style: :thick, color: '000000', edges: [:bottom] },
                                                      alignment: { horizontal: :center, vertical: :center, wrap_text: true })
        green_col_with_thin_border = style.add_style(bg_color: '99CC99', b: true, border: { style: :thin, color: '000000', edges: [:bottom] },
                                                     alignment: { horizontal: :center, vertical: :center, wrap_text: true })
        green_only_col = style.add_style(bg_color: '99CC99',
                                         alignment: { horizontal: :center,
                                                      vertical: :center, wrap_text: true })
        main_text_with_border = style.add_style(b: true,
                                                alignment: { horizontal: :center }, border:  Axlsx::STYLE_THIN_BORDER)
        main_text_with_thick_border = style.add_style(b: true,
                                                      alignment: { horizontal: :center }, border: { style: :thick, color: '000000', edges: [:bottom] })
        text_color_with_thick_border = style.add_style(b: true, fg_color: 'FF0000',
                                                       alignment: { horizontal: :center }, border: { style: :thick, color: '000000', edges: [:bottom] })
        text_color_with_border = style.add_style(b: true, fg_color: 'FF0000',
                                                 alignment: { horizontal: :center }, border: Axlsx::STYLE_THIN_BORDER)
        border_row = style.add_style(border: { style: :thin, color: '000000', edges: [:bottom] })
        last_row = style.add_style(border: { style: :thick, color: '000000', edges: [:bottom] })

        wb.add_worksheet(name: 'Report') do |sheet|
          sheet.add_row ['Current Month - Mar 22'], style: [main_text]
          sheet.add_row ['MTD CLOSED', '', '', '', '', 'Comm', 'Comm'],
                        style: [grey_col, green_col, green_col, green_col, green_col,
                                green_col, pink_col, pink_col, pink_col], height: 14.3
          sheet.merge_cells 'A2:E2'
          sheet.merge_cells 'G2:I2'
          sheet.merge_cells 'F2:F4'
          sheet.merge_cells 'A3:E3'
          sheet.merge_cells 'G3:I3'
          sheet.merge_cells 'G3:G4'
          sheet.merge_cells 'H3:H4'
          sheet.merge_cells 'I3:I4'
          sheet.merge_cells 'F6:F9'
          sheet.add_row ['', '', '', '', '', '', 'Released Before', 'Released Now', 'On Hold'],
                        style: [grey_col, green_col, green_col, green_col, green_col,
                                green_col, pink_col, pink_red_col, pink_col], height: 14.3
          sheet.add_row ['Company Name', 'Invoice no.', 'Payment Mode', 'Payment Date', 'Y/N', '', '', '', ''],
                        style: [yellow_col, yellow_col, yellow_col, yellow_col, yellow_col,
                                green_col, pink_col, pink_red_col, pink_col], height: 50
          sheet.add_row ['', '', '', '', '', 'S$', 'S$', 'S$', 'S$'],
                        style: [border_row, border_row, border_row, border_row, border_row, green_col_with_thin_border,
                                main_text_with_border, text_color_with_border, main_text_with_border]
          sheet.add_row ['', '', '', '', '', '', '', '', ''], style: [nil, nil, nil, nil, nil, green_col, nil, nil, nil]
          @reports.each do |report|
            sheet.add_row [report.company_name, report.invoice_no, report.payment_mode, report.payment_date, report.y_n,
                           report.comm, report.released_before, report.released_now, report.on_hold],
                          style: [left_alignment, left_alignment, center_alignment, center_alignment,
                                  center_alignment, green_only_col, center_alignment, text_color, center_alignment]
          end
          sheet.add_row ['', '', '', '', '', '', '', '', ''],  style: [border_row, border_row, border_row, border_row, border_row, green_col_with_thin_border, border_row, border_row,
                                                                       border_row]
          sheet.add_row ['', '', '', '', '', 'S$', 'S$', 'S$', 'S$'],
                        style: [last_row, last_row, last_row, last_row, last_row, green_col_with_thick_border,
                                main_text_with_thick_border, text_color_with_thick_border, main_text_with_thick_border]
          sheet.add_row ['', '', '', '', '', '', '', '', '$    -']
          sheet.add_row []
          sheet.add_row []
          sheet.add_row []

          # next tables starting

          sheet.add_row ['Outstanding Month'], style: [main_text]
          sheet.add_row ['MTD CLOSED', '', '', '', '', 'Comm', 'Comm'],
                        style: [grey_col, green_col, green_col, green_col, green_col,
                                green_col, pink_col, pink_col, pink_col], height: 14.3
          sheet.merge_cells 'A16:E16'
          sheet.merge_cells 'G16:I16'
          sheet.merge_cells 'F16:F18'
          sheet.merge_cells 'A17:E17'
          sheet.merge_cells 'G17:G18'
          sheet.merge_cells 'H17:H18'
          sheet.merge_cells 'I17:I18'
          sheet.merge_cells 'F20:F23'
          sheet.add_row ['', '', '', '', '', '', 'Released Before', 'Released Now', 'On Hold'],
                        style: [grey_col, green_col, green_col, green_col, green_col,
                                green_col, pink_col, pink_red_col, pink_col], height: 14.3
          sheet.add_row ['Company Name', 'Invoice no.', 'Payment Mode', 'Payment Date', 'Y/N', '', '', '', ''],
                        style: [yellow_col, yellow_col, yellow_col, yellow_col, yellow_col,
                                green_col, pink_col, pink_red_col, pink_col], height: 50
          sheet.add_row ['', '', '', '', '', 'S$', 'S$', 'S$', 'S$'], style: [border_row, border_row, border_row, border_row, border_row, green_col_with_thin_border,
                                                                              main_text_with_border, text_color_with_border, main_text_with_border], height: 30
          sheet.add_row ['', '', '', '', '', '', '', '', ''],
                        style: [nil, nil, nil, nil, nil, green_col,
                                nil, nil, nil]
          @reports.each do |report|
            sheet.add_row [report.company_name, report.invoice_no, report.payment_mode, report.payment_date, report.y_n,
                           report.comm, report.released_before, report.released_now, report.on_hold],
                          style: [left_alignment, left_alignment, center_alignment, center_alignment,
                                  center_alignment, green_col, center_alignment, text_color, center_alignment]
          end
          sheet.add_row ['', '', '', '', '', '', '', '', ''],
                        style: [border_row, border_row, border_row, border_row, border_row, green_col_with_thin_border, border_row, border_row,
                                border_row]
          sheet.add_row ['', '', '', '', '', 'S$', 'S$', 'S$', 'S$'],
                        style: [last_row, last_row, last_row, last_row, last_row, green_col_with_thick_border,
                                main_text_with_thick_border, text_color_with_thick_border, main_text_with_thick_border]
          sheet.add_row ['', '', '', '', '', '', '', '', '$    -']

          # last calculation
          sheet.add_row []
          sheet.add_row []
          sheet.add_row []
          sheet.add_row ['', '', '', '', '', 'Comm Release', '', '$ 1,560.00', ''],
                        style: [nil, nil, nil, nil, nil, right_alignment, nil, text_color, nil]
          sheet.add_row ['', '', '', '', '', 'Sales Awards', '', '$ -   ', ''],
                        style: [nil, nil, nil, nil, nil, right_alignment, nil, text_color, nil]
          sheet.add_row ['', '', '', '', '', 'Adjustments', '', '', ''],
                        style: [nil, nil, nil, nil, nil, right_alignment_with_thin_border, border_row, border_row,
                                border_row]
          sheet.add_row ['', '', '', '', '', 'Comm Release - Mar 22', '', '$ 1,560.00', ''],
                        style: [nil, nil, nil, nil, nil, main_text_with_thick_border, last_row,
                                text_color_with_thick_border, last_row]
        end
      end
      p.serialize('report.xlsx')
    end
  end
  # def table_header

  # end
end

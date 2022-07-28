class AxlsxService
  def initialize(reports)
    @reports = reports
  end

  def create_report
    Axlsx::Package.new do |p|
      wb = p.workbook
      wb.styles do |style|
        border = style.add_style(border: Axlsx::STYLE_THIN_BORDER)
        yellow_col = style.add_style(bg_color: 'CCCC66', b: true, border: Axlsx::STYLE_THIN_BORDER,
                                     alignment: { horizontal: :center, vertical: :center, wrap_text: true })
        green_col = style.add_style(bg_color: '99CC99', b: true, border: Axlsx::STYLE_THIN_BORDER,
                                    alignment: { horizontal: :center, vertical: :center, wrap_text: true })
        pink_red_col = style.add_style(fg_color: 'FF0000', b: true, bg_color: 'FF99FF',
                                       border: Axlsx::STYLE_THIN_BORDER, alignment: { horizontal: :center, vertical: :center, wrap_text: true })
        pink_col = style.add_style(bg_color: 'FF99FF', b: true, border: { style: :thin, color: '000000', edges: %i[bottom top right left] },
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
          row_ft_1 = sheet.add_row ['MTD CLOSED', '', '', '', '', 'Comm', 'Comm', '', ''],
                                   style: [grey_col, green_col, green_col, green_col, green_col,
                                           green_col, pink_col, pink_col, pink_col], height: 14.3
          str_ft_1 = 'MTD CLOSED'
          if str_ft_1 == 'MTD CLOSED'
            row_ft_1.cells.first.value = str_ft_1
            sheet.merge_cells("#{row_ft_1.cells.first.r}:#{row_ft_1.cells[4].r}")
            sheet.merge_cells("#{row_ft_1.cells[6].r}:#{row_ft_1.cells[8].r}")
          end
          str_ft_2 = ''
          str_ft_3 = 'Comm'
          str_ft_4 = 'Released Before'
          str_ft_5 = 'Released Now'
          str_ft_6 = 'On Hold'
          row_ft_2 = sheet.add_row ['', '', '', '', '', '', 'Released Before', 'Released Now', 'On Hold'],
                                   style: [grey_col, green_col, green_col, green_col, green_col,
                                           green_col, pink_col, pink_red_col, pink_col], height: 14.3
          if str_ft_2 == ''
            row_ft_2.cells.first.value = str_ft_2
            sheet.merge_cells("#{row_ft_2.cells[0].r}:#{row_ft_2.cells[4].r}")
          end
          row_ft_3 = sheet.add_row ['Company Name', 'Invoice no.', 'Payment Mode (U/C)', 'Payment Date', 'Y/N', '', '', '', ''],
                                   style: [yellow_col, yellow_col, yellow_col, yellow_col, yellow_col,
                                           green_col, pink_col, pink_red_col, pink_col], height: 50

          sheet.merge_cells("#{row_ft_1.cells[5].r}:#{row_ft_3.cells[5].r}") if str_ft_3 == 'Comm'
          if str_ft_4 == 'Released Before'
            row_ft_2.cells[6].value = str_ft_4
            sheet.merge_cells("#{row_ft_2.cells[6].r}:#{row_ft_3.cells[6].r}")
          end

          if str_ft_5 == 'Released Now'
            row_ft_2.cells[7].value = str_ft_5
            sheet.merge_cells("#{row_ft_2.cells[7].r}:#{row_ft_3.cells[7].r}")
          end

          if str_ft_6 == 'On Hold'
            row_ft_2.cells[8].value = str_ft_6
            sheet.merge_cells("#{row_ft_2.cells[8].r}:#{row_ft_3.cells[8].r}")
          end
          sheet.add_row ['', '', '', '', '', 'S$', 'S$', 'S$', 'S$'],
                        style: [border, border, border, border, border, green_col_with_thin_border,
                                main_text_with_border, text_color_with_border, main_text_with_border]
          sheet.add_row ['', '', '', '', '', '', '', '', ''],
                        style: [nil, nil, nil, nil, nil, green_only_col, nil, nil, nil]

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
          sheet.add_row []

          # next tables starting
          str_st_1 = 'MTD CLOSED'
          sheet.add_row ['Outstanding Month'], style: [main_text]
          row_st_1 = sheet.add_row ['MTD CLOSED', '', '', '', '', 'Comm', 'Comm', '', ''],
                                   style: [grey_col, green_col, green_col, green_col, green_col,
                                           green_col, pink_col, pink_red_col, pink_col], height: 14.3

          if str_st_1 == 'MTD CLOSED'
            row_st_1.cells.first.value = str_st_1
            sheet.merge_cells("#{row_st_1.cells.first.r}:#{row_st_1.cells[4].r}")
            sheet.merge_cells("#{row_st_1.cells[6].r}:#{row_st_1.cells[8].r}")
          end
          row_st_2 = sheet.add_row ['', '', '', '', '', '', 'Released Before', 'Released Now', 'On Hold'],
                                   style: [grey_col, green_col, green_col, green_col, green_col,
                                           green_col, pink_col, pink_red_col, pink_col], height: 14.3
          str_st_2 = ''
          if str_st_2 == ''
            row_st_2.cells.first.value = str_st_2
            sheet.merge_cells("#{row_st_2.cells[0].r}:#{row_st_2.cells[4].r}")
          end
          row_st_3 = sheet.add_row ['Company Name', 'Invoice no.', 'Payment Mode (U/C)', 'Payment Date', 'Y/N', '', '', '', ''],
                                   style: [yellow_col, yellow_col, yellow_col, yellow_col, yellow_col,
                                           green_col, pink_col, pink_red_col, pink_col], height: 50
          str_st_4 = 'Released Before'
          str_st_5 = 'Released Now'
          str_st_6 = 'On Hold'

          if str_st_4 == 'Released Before'
            row_st_2.cells[6].value = str_st_4
            sheet.merge_cells("#{row_st_2.cells[6].r}:#{row_st_3.cells[6].r}")
          end

          if str_st_5 == 'Released Now'
            row_st_2.cells[7].value = str_st_5
            sheet.merge_cells("#{row_st_2.cells[7].r}:#{row_st_3.cells[7].r}")
          end

          if str_st_6 == 'On Hold'
            row_st_2.cells[8].value = str_st_6
            sheet.merge_cells("#{row_st_2.cells[8].r}:#{row_st_3.cells[8].r}")
          end

          str_st_3 = 'Comm'
          sheet.merge_cells("#{row_st_1.cells[5].r}:#{row_st_3.cells[5].r}") if str_st_3 == 'Comm'

          sheet.add_row ['', '', '', '', '', 'S$', 'S$', 'S$', 'S$'], style: [border, border, border, border, border, green_col_with_thin_border,
                                                                              main_text_with_border, text_color_with_border, main_text_with_border], height: 30
          sheet.add_row ['', '', '', '', '', '', '', '', ''],
                        style: [nil, nil, nil, nil, nil, green_only_col,
                                nil, nil, nil]
          @reports.each do |report|
            sheet.add_row [report.company_name, report.invoice_no, report.payment_mode, report.payment_date, report.y_n,
                           report.comm, report.released_before, report.released_now, report.on_hold],
                          style: [left_alignment, left_alignment, center_alignment, center_alignment,
                                  center_alignment, green_only_col, center_alignment, text_color, center_alignment]
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
end

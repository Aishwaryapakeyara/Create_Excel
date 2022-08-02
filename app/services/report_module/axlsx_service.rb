# frozen_string_literal: true

# require_relative 'styling'

module ReportModule
  # app/services/module_1/axlsx_service.rb
  class AxlsxService
    def initialize(reports)
      @reports = reports
    end

    def create_report
      Axlsx::Package.new do |p|
        wb = p.workbook
        wb.styles do |style|
          ServiceModule::Styling.style_call(style)
          # @main_text = style.add_style(b: true, alignment: { horizontal: :center })
          wb.add_worksheet(name: 'Report') do |sheet|
            sheet_creation(sheet, @main_text)
          end
        end
        p.serialize('report.xlsx')
      end
    end

    def sheet_creation(sheet, _style)
      sheet.add_row ['Current Month - Mar 22'], style: [@main_text]
      table_creation(sheet, _style)
      3.times { sheet.add_row [] }
      sheet.add_row ['Outstanding Month'], style: [@main_text]
      table_creation(sheet, _style)
      3.times { sheet.add_row [] }
      file_footer(sheet, _style)
    end

    def table_creation(sheet, _style)
      table_header(sheet, _style)
      table_data(sheet, _style)
      table_footer(sheet, _style)
    end

    def table_header(sheet, _style)
      # ServiceModule::Styling.header_border(style)
      # header_styles(style)
      ServiceModule::Styling.style_call(_style)
      req_values
      row1 = sheet.add_row ['MTD CLOSED', '', '', '', '', 'Comm', 'Comm', '', ''],
                           style: [@grey_col, @green_col, @green_col, @green_col, @green_col,
                                   @green_col, @pink_col, @pink_col, @pink_col], height: 14.3
      five_col_merge(sheet, @str1, row1)
      row2 = sheet.add_row ['', '', '', '', '', '', 'Released Before', 'Released Now', 'On Hold'],
                           style: [@grey_col, @green_col, @green_col, @green_col, @green_col,
                                   @green_col, @pink_col, @pink_red_col, @pink_col], height: 14.3
      three_merge(sheet, '', row2)
      row3 = sheet.add_row ['Company Name', 'Invoice no.', 'Payment Mode (U/C)', 'Payment Date', 'Y/N', '', '', '', ''],
                           style: [@yellow_col, @yellow_col, @yellow_col, @yellow_col, @yellow_col,
                                   @green_col, @pink_col, @pink_red_col, @pink_col], height: 50
      three_row_merge(sheet, @str3, row1, row3)
      two_row_merge(sheet, @str4, row2, row3, 'Released Before', 6)
      two_row_merge(sheet, @str5, row2, row3, 'Released Now', 7)
      two_row_merge(sheet, @str6, row2, row3, 'On Hold', 8)
      sheet.add_row ['', '', '', '', '', 'S$', 'S$', 'S$', 'S$'],
                    style: [@border, @border, @border, @border, @border, @green_col_with_thin_border,
                            @main_text_with_border, @text_color_with_border, @main_text_with_border]
      sheet.add_row ['', '', '', '', '', '', '', '', ''],
                    style: [nil, nil, nil, nil, nil, @green_only_col, nil, nil, nil]
    end

    def three_merge(sheet, str, row)
      return unless str == ''

      row.cells.first.value = str
      sheet.merge_cells("#{row.cells[0].r}:#{row.cells[4].r}")
    end

    def table_footer(sheet, _style)
      ServiceModule::Styling.style_call(_style)
      sheet.add_row ['', '', '', '', '', '', '', '', ''],
                    style: [@border_row, @border_row, @border_row, @border_row, @border_row,
                            @green_col_with_thin_border, @border_row, @border_row, @border_row]
      sheet.add_row ['', '', '', '', '', 'S$', 'S$', 'S$', 'S$'],
                    style: [@last_row, @last_row, @last_row, @last_row, @last_row, @green_col_with_thick_border,
                            @main_text_with_thick_border, @text_color_with_thick_border, @main_text_with_thick_border]
      sheet.add_row ['', '', '', '', '', '', '', '', '$    -']
    end

    def file_footer(sheet, _style)
      # common_footer_style(style)
      # file_footer_style(style)
      ServiceModule::Styling.style_call(_style)
      sheet.add_row ['', '', '', '', '', 'Comm Release', '', '$ 1,560.00', ''],
                    style: [nil, nil, nil, nil, nil, @right_alignment, nil, @text_color, nil]
      sheet.add_row ['', '', '', '', '', 'Sales Awards', '', '$ -   ', ''],
                    style: [nil, nil, nil, nil, nil, @right_alignment, nil, @text_color, nil]
      sheet.add_row ['', '', '', '', '', 'Adjustments', '', '', ''],
                    style: [nil, nil, nil, nil, nil, @right_alignment_with_thin_border, @border_row, @border_row,
                            @border_row]
      sheet.add_row ['', '', '', '', '', 'Comm Release - Mar 22', '', '$ 1,560.00', ''],
                    style: [nil, nil, nil, nil, nil, @main_text_with_thick_border, @last_row,
                            @text_color_with_thick_border, @last_row]
    end

    def five_col_merge(sheet, str, row)
      return unless str == 'MTD CLOSED'

      row_cell = row.cells
      row_cell.first.value = str
      sheet.merge_cells("#{row_cell.first.r}:#{row_cell[4].r}")
      sheet.merge_cells("#{row_cell[6].r}:#{row_cell[8].r}")
    end

    def two_row_merge(sheet, str, row1, row2, str_value, cell_no)
      return unless str == str_value

      row1.cells[cell_no].value = str
      sheet.merge_cells("#{row1.cells[cell_no].r}:#{row2.cells[cell_no].r}")
    end

    def three_row_merge(sheet, str, row1, row2)
      sheet.merge_cells("#{row1.cells[5].r}:#{row2.cells[5].r}") if str == 'Comm'
    end

    def table_data(sheet, _style)
      # table_data_style(style)
      # common_styles(style)
      ServiceModule::Styling.style_call(_style)
      @reports.each do |report|
        sheet.add_row [report.company_name, report.invoice_no, report.payment_mode, report.payment_date, report.y_n,
                       report.comm, report.released_before, report.released_now, report.on_hold],
                      style: [@left_alignment, @left_alignment, @center_alignment, @center_alignment,
                              @center_alignment, @green_only_col, @center_alignment, @text_color, @center_alignment]
      end
    end

    def req_values
      @str1 = 'MTD CLOSED'
      @str2 = ''
      @str3 = 'Comm'
      @str4 = 'Released Before'
      @str5 = 'Released Now'
      @str6 = 'On Hold'
    end
  end
end

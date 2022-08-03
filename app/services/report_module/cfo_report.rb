# frozen_string_literal: true

module ReportModule
  # app/services/report_module/cfo_report.rb
  class CfoReport
    def initialize
      @people = prepared_hash
    end

    def create_cfo_report
      Axlsx::Package.new do |p|
        wrb = p.workbook
        wrb.styles do |style|
          sheet_style(style)
          workbook_creation(wrb)
        end
        p.serialize('CFO_report.xlsx')
      end
    end

    def workbook_creation(wrb)
      wrb.add_worksheet(name: 'CFO_Report') do |sheet|
        @final_sheet_header = table_header.uniq
        sheet.add_row @final_sheet_header, style: [nil, nil, nil, nil, @right_border]
        @final_sheet_header.shift
        table_data(sheet)
        mid_row_data(sheet)
        table_footer(sheet)
      end
    end

    def table_footer(sheet)
      2.times do
        sheet.add_row ['', '', '', '', '', '', '', '', '', '', ''],
                      style: [nil, nil, nil, nil, @right_border, nil, nil, nil, nil, nil, nil]
      end
      total_row(sheet)
    end

    def table_header
      sheet_header = ['Sales Person']
      @people.each_key do |key|
        sheet_header.concat(@people[key].keys)
      end
      sheet_header
    end

    def mid_row_data(sheet)
      len = @final_sheet_header.length
      mid_row = ['Sales Person D']
      len.times { mid_row.push('') }
      mid_row_style = []
      4.times { mid_row_style.push(@top_border) }
      mid_row_style.push(@right_top_border)
      (len - 4).times { mid_row_style.push(@top_border) }
      sheet.add_row mid_row, style: mid_row_style
    end

    def total_row(sheet)
      sheet.add_row ['Total', '', column_sum('RMS New'), column_sum('RMS Renewal'), column_sum('SCCB New'),
                     column_sum('Duntrade'), column_sum('CS Subscription'), column_sum('CM/TM'), column_sum('Seminar'),
                     column_sum('S&MS'), column_sum('D&B Hoovers')],
                    style: [@bold_text_with_l_align, nil, @bold_text_with_r_align, @bold_text_with_r_align,
                            @bold_text_with_r_align, @b_text_r_align_border, @bold_text_with_r_align,
                            @bold_text_with_r_align, @bold_text_with_r_align, @bold_text_with_r_align,
                            @bold_text_with_r_align, @bold_text_with_r_align]
    end

    def table_data(sheet)
      @people.each_key do |key|
        arr = [key]
        @final_sheet_header.each_index do |i|
          arr.push(@people[key][@final_sheet_header[i]])
        end
        @style_arr = [nil, nil]
        @style_arr.push(@right_alignment, @right_alignment, @right_alignment_with_right_border)
        (arr.length - 5).times { @style_arr.push(@right_alignment) }
        sheet.add_row arr, style: @style_arr
      end
    end

    def column_sum(col)
      @people.values.map { |a| a[col] }.sum
    end

    def sheet_style(style)
      @top_border = style.add_style(border: { style: :thin, color: '000000', edges: [:top] })
      @right_border = style.add_style(border: { style: :thin, color: '000000', edges: [:right] })
      @right_top_border = style.add_style(border: { style: :thin, color: '000000', edges: %i[right top] })
      @bold_text_with_r_align = style.add_style(b: true, alignment: { horizontal: :right })
      @bold_text_with_l_align = style.add_style(b: true, alignment: { horizontal: :left })
      @right_alignment = style.add_style(alignment: { horizontal: :right })
      @b_text_r_align_border = style.add_style(b: true, alignment: { horizontal: :right },
                                               border: { style: :thin, color: '000000', edges: [:left] })
      @right_alignment_with_right_border = style.add_style(alignment: { horizontal: :right },
                                                           border: { style: :thin, color: '000000', edges: [:right] })
    end

    def prepared_hash
      @hash = { 'Sales Person A' => { 'Team/Group' => 'RMS', 'RMS New' => 200, 'RMS Renewal' => 300, 'SCCB New' => 50,
                                      'Duntrade' => 0, 'CS Subscription' => 0, 'CM/TM' => 0, 'Seminar' => 50,
                                      'S&MS' => 0, 'D&B Hoovers' => 0 },
                'Sales Person B' => { 'Team/Group' => 'SCCB New', 'RMS New' => 200, 'RMS Renewal' => 0, 'SCCB New' => 600,
                                      'Duntrade' => 100, 'CS Subscription' => 0, 'CM/TM' => 0, 'Seminar' => 0,
                                      'S&MS' => 500, 'D&B Hoovers' => 0 },
                'Sales Person C' => { 'Team/Group' => 'RMS', 'RMS New' => 0, 'RMS Renewal' => 0, 'SCCB New' => 0,
                                      'Duntrade' => 0, 'CS Subscription' => 0, 'CM/TM' => 0, 'Seminar' => 0,
                                      'S&MS' => 500, 'D&B Hoovers' => 200 } }
    end
  end
end

module ReportModule
  # app/services/report_module/cfo_report.rb
  class CfoReport
    def initialize
      @people = prepared_hash
    end

    def create_cfo_report
      Axlsx::Package.new do |p|
        wb = p.workbook
        wb.styles do |style|
          sheet_style(style)
          wb.add_worksheet(name: 'CFO_Report') do |sheet|
            sheet_header = ['Sales Person']
            @people.each_key do |key|
              sheet_header.concat(@people[key].keys)
            end


            # sheet.add_row ['Sales Person', 'Team/Group', 'RMS New', 'RMS Renewal', 'SCCB New New', 'Duntrade',
            #                'CS Subscription', 'CM/TM', 'Seminar', 'S&MS', 'D&B Hoovers'], style: [nil, nil, nil, nil, @right_border, nil, nil, nil, nil, nil, nil]
            sheet.add_row sheet_header.uniq
            # @persons.each_pair do |person|  
            #   sheet.add_row [@persons.key[0], person[key][:Team], person[key][:RMS New], person[key][:RMS Renewal], person[key][:SCCB New],
            #                  person[key][:Duntrade], person[key][:CS Subscription], person[key][:CM/TM], person[key][:Seminar], person[key][:S&MS], person[key][:D&B Hoovers]],
            #                 style: [nil, nil, @right_alignment, @right_alignment, @right_alignment_with_right_border, @right_alignment,
            #                         @right_alignment, @right_alignment, @right_alignment, @right_alignment, @right_alignment]
            
            sheet_header.uniq.shift()
            arr = []
            @people.each_key do |key|
              sheet_header.each_index do |i|
                while i >= 1 && i <= 10
                  arr.push(@people[key][sheet_header[i]])
                end
              end
              sheet.add_row arr
            end


          #   @people.each_key do |key|
          #     arr.unshift(key).uniq
          #   #   sheet.add_row [key, @people[key][sheet_header[1]], @people[key][sheet_header[2]], @people[key]['RMS Renewal'], @people[key]['SCCB New'],
          #   #   @people[key]['Duntrade'], @people[key]['CS Subscription'], @people[key]['CM/TM'], @people[key]['Seminar'], @people[key]['S&MS'], @people[key]['D&B Hoovers']]
          # sheet.add_row arr
          # end
            mid_row = sheet.add_row ['Sales Person D', '', '', '', '', '', '', '', '', '', ''],
                                    style: [@top_border, @top_border, @top_border, @top_border, @right_top_border,
                                            @top_border, @top_border, @top_border, @top_border, @top_border, @top_border]
           2.times { sheet.add_row ['', '', '', '', '', '', '', '', '', '', ''],
           style: [nil, nil, nil, nil, @right_border, nil, nil, nil, nil, nil, nil] }
            # sheet.add_row ['', '', '', '', '', '', '', '', '', '', ''],
            #               style: [nil, nil, nil, nil, @right_border, nil, nil, nil, nil, nil, nil]
            # sheet.add_row ['', '', '', '', '', '', '', '', '', '', ''],
            #               style: [nil, nil, nil, nil, @right_border, nil, nil, nil, nil, nil, nil]
            num = mid_row.index
            sheet.add_row ['Total', '', "=SUM(C2:C#{num})", "=SUM(D2:D#{num})", "=SUM(E2:E#{num})", "=SUM(F2:F#{num})", "=SUM(G2:G#{num})", "=SUM(H2:H#{num})", "=SUM(I2:I#{num})", "=SUM(J2:J#{num})", "=SUM(K2:K#{num})"],
                          style: [@bold_text_with_l_align, nil, @bold_text_with_r_align, @bold_text_with_r_align,
                                  @bold_text_with_r_align, @b_text_r_align_border, @bold_text_with_r_align, @bold_text_with_r_align,
                                  @bold_text_with_r_align, @bold_text_with_r_align, @bold_text_with_r_align, @bold_text_with_r_align]
            # sheet.add_row ['Total'], style: [@bold_text_with_l_align]

            # cells = sheet["C#{num}:K#{num}"]
            # cells.each do |cell|
            #   row_index = mid_row.row.index - 1
            #   cell.type = :string
            #   cell.value = "=SUM(C2:B#{row_index})"
            # end
          end
        end
        p.serialize('CFO_report.xlsx')
      end
    end

    def sheet_style(style)
      @top_border = style.add_style(border: { style: :thin, color: '000000', edges: [:top] })
      @right_border = style.add_style(border: { style: :thin, color: '000000', edges: [:right] })
      @right_top_border = style.add_style(border: { style: :thin, color: '000000', edges: %i[right top] })
      @bold_text_with_r_align = style.add_style(b: true,
                                                alignment: { horizontal: :right })
      @bold_text_with_l_align = style.add_style(b: true,
                                                alignment: { horizontal: :left })
      @right_alignment = style.add_style(alignment: { horizontal: :right })
      @b_text_r_align_border = style.add_style(b: true,
                                               alignment: { horizontal: :right }, border: { style: :thin, color: '000000', edges: [:left] })
      @right_alignment_with_right_border = style.add_style(
        alignment: { horizontal: :right }, border: { style: :thin, color: '000000', edges: [:right] }
      )
    end

    def prepared_hash
      @hash = {
        'Sales Person A' => { 'Team/Group' => 'RMS',
                              'RMS New' => 200, 'RMS Renewal' => 300,
                              'SCCB New' => 50,
                              'Duntrade' => 0,
                              'CS Subscription' => 0,
                              'CM/TM' => 0,
                              'Seminar' => 50,
                              'S&MS' => 0,
                              'D&B Hoovers' => 0 },
        'Sales Person B' => { 'Team/Group' => 'SCCB New',
                              'RMS New' => 100,
                              'RMS Renewal' => 0,
                              'SCCB New' => 600,
                              'Duntrade' => 100,
                              'CS Subscription' => 0,
                              'CM/TM' => 0,
                              'Seminar' => 0,
                              'S&MS' => 500,
                              'D&B Hoovers' => 0 },
        'Sales Person C' => { 'Team/Group' => 'RMS',
                              'RMS New' => 0,
                              'RMS New2' => 0,
                              'RMS Renewal' => 0,
                              'SCCB New' => 0,
                              'Duntrade' => 0,
                              'CS Subscription' => 0,
                              'CM/TM' => 0,
                              'Seminar' => 0,
                              'S&MS' => 500,
                              'D&B Hoovers' => 200 }
      }
    end
  end
end

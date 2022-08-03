# frozen_string_literal: true

module ReportModule
  # app/services/srvice_module/styling.rb
  class Styling
    def self.style_call(style)
      @border = style.add_style(border: Axlsx::STYLE_THIN_BORDER)
      @main_text_with_border = style.add_style(b: true,
                                               alignment: { horizontal: :center }, border:  Axlsx::STYLE_THIN_BORDER)
      @text_color_with_border = style.add_style(b: true, fg_color: 'FF0000',
                                                alignment: { horizontal: :center }, border: Axlsx::STYLE_THIN_BORDER)
      @green_col = style.add_style(bg_color: '99CC99', b: true, border: Axlsx::STYLE_THIN_BORDER,
                                   alignment: { horizontal: :center, vertical: :center, wrap_text: true })
      @grey_col = style.add_style(bg_color: 'A9A9A9', b: true, border: Axlsx::STYLE_THIN_BORDER,
                                  alignment: { horizontal: :center, vertical: :center, wrap_text: true })
      @yellow_col = style.add_style(bg_color: 'CCCC66', b: true, border: Axlsx::STYLE_THIN_BORDER,
                                    alignment: { horizontal: :center, vertical: :center, wrap_text: true })

      @pink_red_col = style.add_style(fg_color: 'FF0000', b: true, bg_color: 'FF99FF',
                                      border: Axlsx::STYLE_THIN_BORDER, alignment: { horizontal: :center, vertical: :center, wrap_text: true })
      @pink_col = style.add_style(bg_color: 'FF99FF', b: true, border: { style: :thin, color: '000000', edges: %i[bottom top right left] },
                                  alignment: { horizontal: :center, vertical: :center, wrap_text: true })
      @border_row = style.add_style(border: { style: :thin, color: '000000',
                                              edges: [:bottom] })
      @last_row = style.add_style(border: { style: :thick, color: '000000', edges: [:bottom] })
      @main_text_with_thick_border = style.add_style(b: true,
                                                     alignment: { horizontal: :center }, border: { style: :thick, color: '000000', edges: [:bottom] })
      @text_color_with_thick_border = style.add_style(b: true, fg_color: 'FF0000',
                                                      alignment: { horizontal: :center }, border: { style: :thick, color: '000000', edges: [:bottom] })
      @border_row = style.add_style(border: { style: :thin,
                                              color: '000000', edges: [:bottom] })
      @last_row = style.add_style(border: { style: :thick,
                                            color: '000000', edges: [:bottom] })
      @main_text_with_thick_border = style.add_style(b: true,
                                                     alignment: { horizontal: :center }, border: { style: :thick, color: '000000', edges: [:bottom] })
      @text_color_with_thick_border = style.add_style(b: true, fg_color: 'FF0000',
                                                      alignment: { horizontal: :center }, border: { style: :thick, color: '000000', edges: [:bottom] })
      @right_alignment = style.add_style(alignment: { horizontal: :right })
      @right_alignment_with_thin_border = style.add_style(alignment: { horizontal: :right },
                                                          border: {
                                                            style: :thin, color: '000000', edges: [:bottom]
                                                          })
      @text_color = style.add_style(
        fg_color: 'FF0000', alignment: { horizontal: :center }
      )
      @left_alignment = style.add_style(alignment: { horizontal: :left })
      @center_alignment = style.add_style(alignment: { horizontal: :center })
      @green_only_col = style.add_style(bg_color: '99CC99',
                                        alignment: { horizontal: :center,
                                                     vertical: :center, wrap_text: true })
      @green_col_with_thick_border = style.add_style(bg_color: '99CC99', b: true, border: { style: :thick, color: '000000', edges: [:bottom] },
                                                     alignment: { horizontal: :center, vertical: :center, wrap_text: true })
      @green_col_with_thin_border = style.add_style(bg_color: '99CC99', b: true, border: { style: :thin, color: '000000', edges: [:bottom] },
                                                    alignment: { horizontal: :center, vertical: :center, wrap_text: true })
      @left_alignment = style.add_style(alignment: { horizontal: :left })
      @center_alignment = style.add_style(alignment: { horizontal: :center })
      @green_only_col = style.add_style(bg_color: '99CC99',
                                        alignment: { horizontal: :center,
                                                     vertical: :center, wrap_text: true })
      @text_color = style.add_style(
        fg_color: 'FF0000', alignment: { horizontal: :center }
      )
      @main_text = style.add_style(b: true, alignment: { horizontal: :center })
    end

    # def self.header_border(style)
    #   @border = style.add_style(border: Axlsx::STYLE_THIN_BORDER)
    #   @main_text_with_border = style.add_style(b: true,
    #                                            alignment: { horizontal: :center }, border:  Axlsx::STYLE_THIN_BORDER)
    #   @text_color_with_border = style.add_style(b: true, fg_color: 'FF0000',
    #                                             alignment: { horizontal: :center }, border: Axlsx::STYLE_THIN_BORDER)
    # end

    #     def self.header_styles(style)
    #       @green_col = style.add_style(bg_color: '99CC99', b: true, border: Axlsx::STYLE_THIN_BORDER,
    #                                    alignment: { horizontal: :center, vertical: :center, wrap_text: true })
    #       @grey_col = style.add_style(bg_color: 'A9A9A9', b: true, border: Axlsx::STYLE_THIN_BORDER,
    #                                   alignment: { horizontal: :center, vertical: :center, wrap_text: true })
    #       @yellow_col = style.add_style(bg_color: 'CCCC66', b: true, border: Axlsx::STYLE_THIN_BORDER,
    #                                     alignment: { horizontal: :center, vertical: :center, wrap_text: true })

    #       @pink_red_col = style.add_style(fg_color: 'FF0000', b: true, bg_color: 'FF99FF',
    #                                       border: Axlsx::STYLE_THIN_BORDER, alignment: { horizontal: :center, vertical: :center, wrap_text: true })
    #       @pink_col = style.add_style(bg_color: 'FF99FF', b: true, border: { style: :thin, color: '000000', edges: %i[bottom top right left] },
    #                                   alignment: { horizontal: :center, vertical: :center, wrap_text: true })
    #     end

    #     def self.common_footer_style(style)
    #       # @border_row = style.add_style(border: { style: :thin, color: '000000', edges: [:bottom] })
    #       # @last_row = style.add_style(border: { style: :thick, color: '000000', edges: [:bottom] })
    #       # @main_text_with_thick_border = style.add_style(b: true,
    #       #                                                alignment: { horizontal: :center }, border: { style: :thick, color: '000000', edges: [:bottom] })
    #       # @text_color_with_thick_border = style.add_style(b: true, fg_color: 'FF0000',
    #       #                                                 alignment: { horizontal: :center }, border: { style: :thick, color: '000000', edges: [:bottom] })
    #     end

    #     def self.file_footer_style(style)
    #       # @right_alignment = style.add_style(alignment: { horizontal: :right })
    #       # @right_alignment_with_thin_border = style.add_style(alignment: { horizontal: :right },
    #       #                                                     border: { style: :thin, color: '000000', edges: [:bottom] })
    #     end

    #     def self.table_footer_style(style)
    #       # @green_col_with_thick_border = style.add_style(bg_color: '99CC99', b: true, border: { style: :thick, color: '000000', edges: [:bottom] },
    #       #                                                alignment: { horizontal: :center, vertical: :center, wrap_text: true })
    #       # @green_col_with_thin_border = style.add_style(bg_color: '99CC99', b: true, border: { style: :thin, color: '000000', edges: [:bottom] },
    #       #                                               alignment: { horizontal: :center, vertical: :center, wrap_text: true })
    #     end

    #     def self.table_data_style(style)
    #       # @left_alignment = style.add_style(alignment: { horizontal: :left })
    #       # @center_alignment = style.add_style(alignment: { horizontal: :center })
    #       # @green_only_col = style.add_style(bg_color: '99CC99',
    #       #                                   alignment: { horizontal: :center,
    #       #                                                vertical: :center, wrap_text: true })
    #     end

    #     def self.common_styles(style)
    #       # @text_color = style.add_style(fg_color: 'FF0000', alignment: { horizontal: :center })
    #       # @main_text = style.add_style(b: true, alignment: { horizontal: :center })
    #     end
  end
end

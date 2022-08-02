class ReportsController < ApplicationController
  def index
    @reports = Report.all
    AxlsxService.new(@reports).create_report
    # respond_to do |format|
    #   format.html
    #   format.xlsx
    # end
    # render xlsx: 'index', filename: 'reports.xlsx'
  end
end

class ReportsController < ApplicationController
  def index
    @reports = Report.all
    # AxlsxService.create(@reports)
    respond_to do |format|
      format.html
      format.xlsx
    end
  end
end

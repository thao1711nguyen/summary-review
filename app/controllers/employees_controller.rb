class EmployeesController < ApplicationController
  def summary
    zipped_file = params[:file]
    if zipped_file && zipped_file.content_type == 'application/zip'
      result_file = Employee.generate_summary(zipped_file)
    end
  end
end

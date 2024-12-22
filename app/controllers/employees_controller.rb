class EmployeesController < ApplicationController
  def new
  end
  def create
    file = params[:file]
    if file.content_type != "application/json"
      flash[:alert] = "You need to submit a json file!"
    else
      employee_list = {}
      begin
        employee_list = JSON.parse(file.read)
      rescue => e
        Rails.logger.error e.message
        redirect_to new_employee_path
      end
      result = Employee.create_employees(employee_list)
      if result.length == 0
        flash[:notice] = "successfully create employees"
        redirect_to summary_reviews_path
      else
        flash[:alert] = result
        redirect_to new_employee_path
      end
    end
  end
  def new_summary
  end
  def summary
    zipped_file = params[:file]
    if zipped_file.content_type == "application/zip"
      Employee.generate_summary(zipped_file)
      file_name = "result.zip"
      zipped_file_path = Rails.root.join("app", "views", "employees", "review", file_name)
      send_file zipped_file_path,
                type: "application/zip",
                filename: file_name,
                diposition: "attachment"

    else
      flash[:alert] = "You need to send a zip file"
      redirect_to summary_reviews_path
    end
  end
end

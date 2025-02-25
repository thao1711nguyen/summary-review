class EmployeesController < ApplicationController
  skip_before_action :verify_authenticity_token, only: :summary if Rails.env.test?
  def new_summary
  end
  def summary
    zipped_file = params[:file]
    if zipped_file.content_type == "application/zip" || zipped_file.content_type == "application/x-zip-compressed"
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

class EmployeesController < ApplicationController
  skip_before_action :verify_authenticity_token, only: :summary if Rails.env.test?
  def new_summary
  end
  def summary
    zipped_file = params[:file]
    if zipped_file.content_type == "application/zip" || zipped_file.content_type == "application/x-zip-compressed"
      zipped_file_path = Employee.generate_summary(zipped_file)
      send_file zipped_file_path,
                type: "application/zip",
                filename: file_name,
                disposition: "attachment"
    else
      flash[:alert] = "You need to send a zip file"
      redirect_to summary_reviews_path
    end
  end  
end

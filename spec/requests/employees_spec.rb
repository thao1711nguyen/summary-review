require 'rails_helper'

RSpec.describe "Employees", type: :request do
  describe "post /employees/summary_reviews" do
    it "return zipped file" do
      file = fixture_file_upload(Rails.root.join(), 'application/x-zip-compressed')
      post summary_reviews_path, 
            params: {
              file: file
            }
      expect(response).to have_http_status(200  )
    end
  end
end

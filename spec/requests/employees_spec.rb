require 'rails_helper'

RSpec.describe "Employees", type: :request do
  describe "post /employees/summary_reviews" do
    it "return zipped file" do
      file = fixture_file_upload(Rails.root.join('test', 'fixtures', 'files', 'NguyenHuynhThanhThao-Review cheo.Dinh tinh.zip'), 'application/x-zip-compressed')
      post "/employees/summary_reviews", 
            params: {
              file: file
            }
      binding.pry
      expect(response).to have_http_status(200)
    end
  end
end

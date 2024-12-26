require 'rails_helper'

RSpec.describe "Employees", type: :request do
  describe "post /employees/summary_reviews" do
    it "return zipped file" do
      post summary_reviews_path, 
            params: {
              file: file
            }
    end
  end
end

class Employee < ApplicationRecord
  has_one_attached :summary_file 
  enum :category, {
    manager: 1, 
    leader: 2, 
    normal: 3
  }
  def self.generate_summary(zipped_file)
    # unzip file 
    # 
  end
end


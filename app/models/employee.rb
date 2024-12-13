class Employee < ApplicationRecord
  has_one_attached :summary_file 
  enum :category, {
    manager: 1, 
    leader: 2, 
    normal: 3
  }
end


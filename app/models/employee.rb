class Employee < ApplicationRecord
  has_one_attached :summary_file
  enum :category, {
    manager: 1,
    leader: 2,
    normal: 3
  }
  validates :name, presence: true
  def self.generate_summary(zipped_file)
    # unzip file
    #
  end
  def create_employees(list)
    errors= []
    if managers = list["manager"]
      managers.each do |manager|
        new_record = Employee.new(name: manager["name"], category: 1)
        unless new_record.save
          errors << new_record.errors.full_messages
        end
      end
    end
    if leaders = list["leader"]
      leaders.each do |leader|
        new_record = Employee.new(name: leader["name"], category: 2)
        unless new_record.save
          errors << new_record.errors.full_messages
        end
      end
    end
    if normals = list["normal"]
      normals.each do |normal|
        new_record = Employee.new(name: normal["name"])
        unless new_record.save
          errors << new_record.errors.full_messages
        end
      end
    end
    errors
  end
end

class Employee < ApplicationRecord
  has_one_attached :summary_file
  enum :category, {
    manager: 1,
    leader: 2,
    normal: 3
  }
  validates :name, presence: true
  validates :full_name, presence: true

  class << self 
    def generate_summary(zipped_file)
      errors = []
      begin 
        # unzip file
        Zip::File.open(zipped_file.path) do |zipfile|
          # iterate through each file 
          zipfile.each do |entry|
            # check if each file is xlsx 
            #if not return the file name
            unless excel_mime_types.include? entry.content_type
              errors << "file [#{entry.name}] is not an excel file"
            else
              owner = get_owner(entry.name)
              unless owner 
               logger.error "can't find the file's owner!"
               return 
              end
              #iterate through each sheet 
               workbook = RubyXL::Parser.parse(entry)
               #sheet tu duy 
              if workbook[0].sheet_name == '考え方- Tu duy'

                analyze_sheet_0(workbook[0], owner)
              else 
                logger.error 'sheet tu duy, wrong name!' 
                return 
              end
                  
               #sheet nhiet tinh 
               #sheet vai tro - chung 
               #sheet vai tro - leader
               #sheet vai tro - manager
            end
          end
        end
      rescue => e 
        logger.error e.message 
      end

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
  private 
  def excel_mime_types 
    [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', # .xlsx
      'application/vnd.ms-excel' # .xls
    ]
  end
  def analyze_sheet_0(owner, sheet)

  end
  def get_owner(file_name)
    Employee.find_by(full_name: file_name.split("-").first)
  end
  def get_employees(names)
    Employee.where("name IN (?)", names)
  end
  def get_headers(sheet)
    sheet[1][3..-1].map do |cell|
      cell.value.split("-").strip
    end
  end
end

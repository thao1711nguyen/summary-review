class Employee < ApplicationRecord
  has_one_attached :summary_file
  scope :managers, lambda { where(category: 1) }
  scope :leaders, lambda { where(category: 2) }
  scope :normals, lambda { where(category: 3) }
  enum :category, {
    manager: 1,
    leader: 2,
    normal: 3
  }
  validates :name, presence: true
  START_ROW = 2

  class << self
    def generate_summary(zipped_file)
      errors = analyze_data(zipped_file)
      return errors if errors.length > 0
      clean_up(Rails.root.join("app", "view", "employees", "review"))
      generate_zip_file
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
    private
    def is_excel_file?(name)
      extension = name.split(".").last.strip
      ['xlsx', 'xls'].include? extension
    end
    def analyze_sheet(owner, sheet, sheet_edges, options={})
      employees = get_employees(sheet, sheet_edges)
      start_row = sheet_edges[:start_row]
      end_row = sheet_edges[:end_row]
      start_col = sheet_edges[:start_col]
      end_col = sheet_edges[:end_col]
      property = sheet.sheet_name.split("-").last.strip
      exclude_list = options[:exclude] ? options[:exclude] : []
      sheet.sheet_data.rows.each.with_index(start_row) do |row, row_index|
        break if row_index > end_row
        row.cells.each.with_index(start_col) do |cell, cell_index|
          break if cell_index > end_col
          i = cell_index - start_col
          next if owner.name == employees[i].name
          next if exclude_list.include? employees[i].name
          employee = employees[i]
          if employee[property]
            begin
              employee[property] = MessagePack.unpack(employee[property])
            rescue => e
              logger.error e.message
            end
          else
            employee[property] = Array.new(end_col-start_col+1) { [] }
          end
          employee[property][i][row_index-start_row] << cell.value
          begin
            employee[property] = employee[property].to_msgpack
            p employee[property]
          rescue => e
            logger.error e.message
          end
          unless employee.save
            logger.error employee.errors.full_messages.join(", ")
          end
        end
      end
    end
  
    def get_owner(file_name)
      Employee.find_by(name: file_name.split("-").strip.first)
    end
    def get_employees(sheet, sheet_edges)
      names = []
      sheet.sheet_data.rows[sheet_edges[:start_row] - 1].cells
      .each.with_index(sheet_edges[:start_col]) do |cell, index|
          name = cell.value.split("-").last.strip 
          names << name
          break if index == sheet_edges[:end_col]
      end
      Employee.where(name: names)
    end
    
    def get_start_end_col(sheet)
      header_row = START_ROW - 1
      result = {}
      sheet.sheet_data.rows[header_row].each_with_index do |column_name, index|
        result[:start_col] = index if column_name.value.include?("レビュー")
        if column_name.value == "レビューReview"
          result[:end_col] = index - 1
          return result
        end
      end
      result[:end_col] ||= sheet.sheet_data.rows[header_row].cells.size - 1
      result
    end
    def get_end_row(sheet, start_col)
      result = {}
      sheet.sheet_data.rows.each_with_index do |row, index|
        if row[start_col] == 0
          result[:end_row] = index - 1
          return result
        end
      end
    end
    
    def sheet_edges_original(sheet)
      result = {
        start_row: START_ROW
      }
      result.merge(get_start_end_col(sheet), get_end_row(sheet))
      # {
      #   "tu_duy" => {
      #     start_row: 2,
      #     end_row: 26,
      #     start_col: 3,
      #     end_col: 19,
      #     property: "tu_duy"
      #   },
      #   "nhiet_tinh" => {
      #     start_row: 2,
      #     end_row: 24,
      #     start_col: 3,
      #     end_col: 19,
      #     property: "nhiet_tinh"
      #   },
      #   "vai_tro" => {
      #     "chung" => {
      #       start_row: 2,
      #       end_row: 21,
      #       start_col: 4,
      #       end_col: 20,
      #       property: "vai_tro"
      #     },
      #     "leader" => {
      #       start_row: 2,
      #       end_row: 21,
      #       start_col: 4,
      #       end_col: 5,
      #       property: "vai_tro"
      #     },
      #     "manager" => {
      #       start_row: 2,
      #       end_row: 21,
      #       start_col: 4,
      #       end_col: 5,
      #       property: "vai_tro"
      #     }
      #   }
      # }
    end
    def sheet_options_result
      {
        "tu_duy" => {
          start_row: 2,
          end_row: 26,
          start_col: 3,
          end_col: 19,
          property: "tu_duy"
        },
        "nhiet_tinh" => {
          start_row: 2,
          end_row: 24,
          start_col: 3,
          end_col: 19,
          property: "nhiet_tinh"
        },
        "vai_tro" => {
          start_row: 2,
          end_row: 21,
          start_col: 4,
          end_col: 20,
          property: "vai_tro"
        }
      }
    end
    def analyze_data(zipped_file)
      errors = []
      clean_up(Rails.root.join('test', 'fixtures', 'files', 'extracted'))
      # begin
        # unzip file
        Zip::File.open(zipped_file.path) do |zipfile|
          # iterate through each file
          
          zipfile.each do |entry|    
            # Extract to file or directory based on name in the archive
            entry.extract(Rails.root.join('test', 'fixtures', 'files', 'extracted', entry.name))
            # check if each file is xlsx
            # if not return the file name
            unless is_excel_file?(entry.name)
              errors << "file [#{entry.name}] is not an excel file"
            else
              owner = get_owner(entry.name)
              unless owner
               logger.error "can't find the file's owner!"
               return
              end
              # iterate through each sheet
              workbook = RubyXL::Parser.parse(File.open(Rails.root.join('test', 'fixtures', 'files', 'extracted', entry.name), 'rb'))
              # sheet tu duy
              # sheet nhiet tinh
              # sheet vai tro - chung
              # sheet vai tro - leader
              # sheet vai tro - manager
              options = sheet_options_original
              workbook.worksheets.each do |sheet|
                sheet_edges = sheet_edges_original(sheet)
                options = {}
                if sheet.sheet_name 
                  options[:exclude] = Employee.managers.pluck(:name) + Employee.leaders.pluck(:name)
                end
                analyze_sheet(owner, sheet, sheet_edges, options)
  
                end
              end
  
            end
          end
          
        end
      # rescue => e
      #   logger.error e.message
      # end
      errors
    end
    def generate_zip_file
        employees = Employee.all
        employees.each do |employee|
          workbook = case employee.category
          when 1
            RubyXL::Parser.parse(Rails.root.join("app", "views", "templates", "cheo", "manager.xlsx"))
          when 2
            RubyXL::Parser.parse(Rails.root.join("app", "views", "templates", "cheo", "leader.xlsx"))
          else
            RubyXL::Parser.parse(Rails.root.join("app", "views", "templates", "cheo", "normal.xlsx"))
          end
          workbook.worksheets.each do |sheet|
            sheet_edges = sheet_edges_original(sheet)
            write_to_sheet(employee, sheet, sheet_edges)
          end
          result_file_name = "#{employee.name}-review.xlsx"
          result_file_path = Rails.root.join("app", "views", "employees", "review", result_file_name)
          workbook.write(result_file_path)
          Zip::File.open(Rails.roout.join("app", "views", "employees", "review", "result.zip"), create: true) do |zip|
            zip.add(result_file_name, result_file_path)
          end
        end
    end
    def write_to_sheet(owner, sheet, sheet_edges)
      employee_names = get_employees(sheet, sheet_edges).pluck(:name)
  
      start_row = sheet_edges[:start_row]
      end_row = sheet_edges[:end_row]
      start_col = sheet_edges[:start_col]
      end_col = sheet_edges[:end_col]
      property = sheet.sheet_name.split("-").last.strip
      sheet.sheet_data.rows.each.with_index(start_row) do |row, row_index|
        break if row_index > end_row
        row.cells.each.with_index do |cell, cell_index|
          break if cell_index > end_col
          i = cell_index - start_col
          next if owner.name == employee_names[i]
          begin
            data = MessagePack.unpack(owner[property])
          rescue => e
            logger.error e.message
          end
          cell[cell_index] = data[i][row_index-start_row]
        end
      end
    end
    def clean_up(folder_name)
      if Dir.exist? folder_name
        Dir.foreach(folder_name) do |file|
          next if file == "." || file == ".." # skip special entries
          file_path = File.join(folder_name, file)
          File.delete(file_path) if File.file?(file_path) # delete only files, not subdirectories
        end
      end
    end
  end
end

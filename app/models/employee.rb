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
    def analyze_sheet(owner, sheet, options)
      headers = get_headers(sheet)
      employees = get_employees(headers)
      employees.each do |employee|
        p employee
      end
      start_row = options[:start_row]
      end_row = options[:end_row]
      start_col = options[:start_col]
      end_col = options[:end_col]
      property = options[:property]
      exclude_list = options[:exclude] ? options[:exclude] : []
      sheet.sheet_data.rows.each_with_index do |row, row_index|
        next if row_index < start_row
        break if row_index > end_row
        row.cells.each_with_index do |cell, cell_index|
          next if cell_index < start_col
          break if cell_index > end_col
          i = cell_index - start_col
          next if owner.name == headers[i]
          next if exclude_list.include? headers[i]
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
          employee[property][i] << cell.value
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
    def get_employees(names)
      p names
      Employee.where(name: names)
    end
    def get_headers(sheet)
      sheet[1][3..19].map do |cell|
        p cell.value.split("-").last.strip
        cell.value.split("-").last.strip
      end
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
                analyze_sheet(owner, sheet, sheet_edges)
                case sheet.sheet_name
                when "考え方- Tu duy"
                  analyze_sheet(owner, sheet, sheet_edges, property: "tu_duy")
                when "熱意- Nhiet tinh"
                  analyze_sheet(owner, sheet, options["nhiet_tinh"])
                when "Vai tro -Chung"
                  chung_options = options["vai_tro"]["chung"]
                  chung_options[:exclude] = Employee.managers.pluck(:name) + Employee.leaders.pluck(:name)
                  analyze_sheet(owner, sheet, chung_options)
                when "Vai tro -Leader"
                  analyze_sheet(owner, sheet, options["vai_tro"]["leader"])
                when "Vai tro -Manager"
                  analyze_sheet(owner, sheet, options["vai_tro"]["manager"])
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
          options = sheet_options_result
          workbook.worksheets.each do |sheet|
            case sheet.sheet_name
            when "考え方- Tu duy"
              write_to_sheet(employee, sheet, options["tu_duy"])
            when "熱意- Nhiet tinh"
              write_to_sheet(employee, sheet, options["nhiet_tinh"])
            else
              write_to_sheet(employee, sheet, options["vai_tro"])
            end
          end
          result_file_name = "#{employee.name}-review.xlsx"
          result_file_path = Rails.root.join("app", "views", "employees", "review", result_file_name)
          workbook.write(result_file_path)
          Zip::File.open(Rails.roout.join("app", "views", "employees", "review", "result.zip"), create: true) do |zip|
            zip.add(result_file_name, result_file_path)
          end
        end
    end
    def write_to_sheet(owner, sheet, options)
      headers = get_headers(sheet)
  
      start_row = options[:start_row]
      end_row = options[:end_row]
      start_col = options[:start_col]
      end_col = options[:end_col]
      property = options[:property]
      sheet.sheet_data.rows.each_with_index do |row, row_index|
        next if row_index < start_row
        break if row_index > end_row
        row.cells.each_with_index do |cell, cell_index|
          next if cell_index < start_col
          break if cell_index > end_col
          i = cell_index - start_col
          next if owner.name == headers[i]
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

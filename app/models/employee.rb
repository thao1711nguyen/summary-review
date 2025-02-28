class Employee < ApplicationRecord
  

  START_ROW = 2
  MANAGER = 1
  LEADER = 2
  NORMAL = 3

  

  class << self

    def generate_summary(zipped_file)
      initialize_employees
      errors = analyze_data(zipped_file)
      return errors if errors.length > 0
      clean_up(Rails.root.join("app", "views", "employees", "review"))
      generate_zip_file
    end
    
   
    private
    def initialize_employees
      @employees = {}
    end
    def is_excel_file?(name)
      extension = name.split(".").last.strip
      ['xlsx', 'xls'].include? extension
    end
    def analyze_sheet(owner, sheet, sheet_edges, options={})
      emp_names = get_names(sheet, sheet_edges)
      start_row = sheet_edges[:start_row]
      end_row = sheet_edges[:end_row]
      start_col = sheet_edges[:start_col]
      end_col = sheet_edges[:end_col]
      property = sheet.sheet_name.split("-").last.strip
      exclude_list = options[:exclude] ? options[:exclude] : []
      result = {}
      sheet.sheet_data.rows.each_with_index do |row, row_index|
        next if row_index < start_row
        break if row_index > end_row
        row.cells.each_with_index do |cell, cell_index|
          next if cell_index < start_col
          break if cell_index > end_col
          i = cell_index - start_col
          next if owner == emp_names[i]
          next if exclude_list.include? emp_names[i]
          unless @employees[emp_names[i]][property]
            @employees[emp_names[i]][property] = Array.new
          end
          result[emp_names[i]] = [] unless result[emp_names[i]]
          next unless cell
          result[emp_names[i]] << cell.value
        end
      end  
      @employees.each do |employee, emp_property|
        next unless emp_property[property]
        @employees[employee][property] << result[employee] if result[employee]
      end
    end
    def get_names(sheet, sheet_edges)
      names = []
      sheet.sheet_data.rows[sheet_edges[:start_row] - 1].cells
      .each_with_index do |cell, index|
        next if index < sheet_edges[:start_col]
        name = cell.value.split("-").last.strip 
        names << name
        break if index == sheet_edges[:end_col]
      end
      names
    end

    def setup_employees(sheet, sheet_edges, category=NORMAL)
      sheet.sheet_data.rows[sheet_edges[:start_row] - 1].cells
      .each_with_index do |cell, index|
          next if index < sheet_edges[:start_col]
          name = cell.value.split("-").last.strip 
          @employees[name] = {}
          @employees[name][:category] = category
          break if index == sheet_edges[:end_col]
      end
      
    end
    
    def get_start_end_col(sheet)
      header_row = START_ROW - 1
      result = {}
      sheet.sheet_data.rows[header_row].cells.each_with_index do |column_name, index|
        if column_name&.value&.include?("レビュー") && result[:start_col].nil?
          result[:start_col] = index
        end
        if column_name && column_name.value == "レビュー\nReview"
          result[:end_col] = index  - 1
          return result
        end
      end
      result[:end_col] ||= sheet.sheet_data.rows[header_row].cells.size - 1
      result
    end

    def get_end_row(sheet, start_col)
      result = {}
      sheet.sheet_data.rows.each_with_index do |row, index|
        next if index < start_col
        if row[1].value.to_s.include?("合計")
          result[:end_row] = index - 1
          return result
        end
      end
    end
    
    def sheet_edges_original(sheet)
      result = {
        start_row: START_ROW
      }
      result = result.merge(get_start_end_col(sheet))
      result.merge(get_end_row(sheet, result[:start_col]))
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
          
          zipfile.each_with_index do |entry, index|    
            # Extract to file or directory based on name in the archive
            entry.extract(Rails.root.join('test', 'fixtures', 'files', 'extracted', entry.name))
            # check if each file is xlsx
            # if not return the file name
            unless is_excel_file?(entry.name)
              errors << "file [#{entry.name}] is not an excel file"
            else
              owner = entry.name.split("-").first.strip
              # iterate through each sheet
              workbook = RubyXL::Parser.parse(File.open(Rails.root.join('test', 'fixtures', 'files', 'extracted', entry.name), 'rb'))
              # sheet tu duy
              # sheet nhiet tinh
              # sheet vai tro - chung
              # sheet vai tro - leader
              # sheet vai tro - manager
              workbook.worksheets.each_with_index do |sheet, sheet_index|
                break if sheet_index > 4
                sheet_edges = sheet_edges_original(sheet)
                if index == 0 
                  setup_employees(sheet, sheet_edges) if sheet_index == 0
                  if sheet.sheet_name.include?("leader")
                    leaders = get_names(sheet, sheet_edges) 
                    @employees.each do |employee, property|
                      if leaders.include?(employee)
                        property[:category] = LEADER
                        property["vai_tro"] = nil
                      end
                    end
                  end
                  if sheet.sheet_name.include?("manager") 
                    managers = get_names(sheet, sheet_edges) 
                    @employees.each do |employee, property|
                      if managers.include?(employee)
                        property[:category] =  MANAGER 
                        property["vai_tro"] = nil
                      end
                    end
                  end
                end
                options = {}
                if sheet.sheet_name.include?("chung") 
                  options[:exclude] = @employees.select {|name, property| property[:category] == MANAGER || property[:category] == LEADER  }.keys
                end
                analyze_sheet(owner, sheet, sheet_edges, options)
  
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
      result_dir_path = Rails.root.join("app", "views", "employees", "review")
      clean_up(result_dir_path)
        @employees.each do |employee, property|
          workbook = case property[:category]
          when 1
            RubyXL::Parser.parse(Rails.root.join("app", "views", "templates", "cheo", "manager.xlsx"))
          when 2
            RubyXL::Parser.parse(Rails.root.join("app", "views", "templates", "cheo", "leader.xlsx"))
          else
            RubyXL::Parser.parse(Rails.root.join("app", "views", "templates", "cheo", "normal.xlsx"))
          end
          workbook.worksheets.each_with_index do |sheet, sheet_index|
            break if sheet_index > 2
            sheet_edges = sheet_edges_original(sheet)
            write_to_sheet(employee, sheet, sheet_edges)
          end
          result_file_name = "#{employee}-review.xlsx"
          result_file_path = Rails.root.join(result_dir_path, result_file_name)
          workbook.write(result_file_path)
        end
      # Create ZIP after processing all employees
      Zip::File.open(Rails.root.join("app", "views", "employees", "review", "result.zip"), create: true) do |zip|
        Dir.glob("#{result_dir_path}/*.xlsx").each do |file|
          zip.add(File.basename(file), file)
        end
      end
    end
    def write_to_sheet(owner, sheet, sheet_edges)
      emp_names = get_names(sheet, sheet_edges)
      start_row = sheet_edges[:start_row]
      end_row = sheet_edges[:end_row]
      start_col = sheet_edges[:start_col]
      end_col = sheet_edges[:end_col]
      property = sheet.sheet_name.split("-").last.strip
      
      for row_index in start_row..end_row 
        for cell_index in start_col..end_col
          i = cell_index - start_col 
          if @employees[owner][property]
            if @employees[owner][property][i]
              data = @employees[owner][property][i][row_index - start_row] 
              sheet.add_cell(row_index, cell_index, data)  
            end
          end
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

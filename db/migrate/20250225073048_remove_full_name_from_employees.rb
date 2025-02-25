class RemoveFullNameFromEmployees < ActiveRecord::Migration[7.2]
  def change
    remove_column :employees, :full_name, :string
  end
end

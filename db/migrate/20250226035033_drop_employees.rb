class DropEmployees < ActiveRecord::Migration[7.2]
  def change
    drop_table :employees
  end
end

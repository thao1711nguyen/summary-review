class ChangeColumnEmployees < ActiveRecord::Migration[7.2]
  def change
    change_column :employees, :tu_duy, 'bytea', using: 'tu_duy::bytea' 
    change_column :employees, :nhiet_tinh, 'bytea', using: 'nhiet_tinh::bytea'
    change_column :employees, :vai_tro, 'bytea', using: 'vai_tro::bytea'
  
  end
end

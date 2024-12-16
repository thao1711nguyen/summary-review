class CreateEmployees < ActiveRecord::Migration[7.2]
  def change
    create_table :employees do |t|
      t.integer :category, null: false, default: 3
      t.string :name, null: false
      t.string :full_name, null: false 
      t.string :tu_duy 
      t.string :nhiet_tinh
      t.string :vai_tro
      t.timestamps
    end
  end
end

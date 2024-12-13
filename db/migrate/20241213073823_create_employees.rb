class CreateEmployees < ActiveRecord::Migration[7.2]
  def change
    create_table :employees do |t|
      t.integer :category, null: false, default: 3
      t.timestamps
    end
  end
end

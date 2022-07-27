class CreateReports < ActiveRecord::Migration[7.0]
  def change
    create_table :reports do |t|
      t.string :company_name
      t.string :invoice_no
      t.string :payment_mode
      t.datetime :payment_date
      t.string :y_n
      t.string :comm
      t.string :released_before
      t.string :released_now
      t.string :on_hold

      t.timestamps
    end
  end
end

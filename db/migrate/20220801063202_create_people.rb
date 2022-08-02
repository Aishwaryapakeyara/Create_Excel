class CreatePeople < ActiveRecord::Migration[7.0]
  def change
    create_table :people do |t|
      t.string :sales_person
      t.string :team
      t.string :rms_new
      t.string :rms_renewal
      t.string :sccb
      t.string :duntrade
      t.string :cs_sub
      t.string :cm_tm
      t.string :seminar
      t.string :s_ms
      t.string :dnb_hovers

      t.timestamps
    end
  end
end


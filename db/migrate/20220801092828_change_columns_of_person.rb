# db/migrate/20220801092828_change_columns_of_person.rb
class ChangeColumnsOfPerson < ActiveRecord::Migration[7.0]
  def change
    change_column :people, :rms_new, :integer, using: 'rms_new::integer'
    change_column :people, :rms_renewal, :integer, using: 'rms_renewal::integer'
    change_column :people, :sccb, :integer, using: 'sccb::integer'
    change_column :people, :duntrade, :integer, using: 'duntrade::integer'
    change_column :people, :cs_sub, :integer, using: 'cs_sub::integer'
    change_column :people, :cm_tm, :integer, using: 'cm_tm::integer'
    change_column :people, :seminar, :integer, using: 'seminar::integer'
    change_column :people, :s_ms, :integer, using: 's_ms::integer'
    change_column :people, :dnb_hovers, :integer, using: 'dnb_hovers::integer'
  end
end

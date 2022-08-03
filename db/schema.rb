# This file is auto-generated from the current state of the database. Instead
# of editing this file, please use the migrations feature of Active Record to
# incrementally modify your database, and then regenerate this schema definition.
#
# This file is the source Rails uses to define your schema when running `bin/rails
# db:schema:load`. When creating a new database, `bin/rails db:schema:load` tends to
# be faster and is potentially less error prone than running all of your
# migrations from scratch. Old migrations may fail to apply correctly if those
# migrations use external dependencies or application code.
#
# It's strongly recommended that you check this file into your version control system.

ActiveRecord::Schema[7.0].define(version: 2022_08_01_092828) do
  # These are extensions that must be enabled in order to support this database
  enable_extension "plpgsql"

  create_table "people", force: :cascade do |t|
    t.string "sales_person"
    t.string "team"
    t.integer "rms_new"
    t.integer "rms_renewal"
    t.integer "sccb"
    t.integer "duntrade"
    t.integer "cs_sub"
    t.integer "cm_tm"
    t.integer "seminar"
    t.integer "s_ms"
    t.integer "dnb_hovers"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  create_table "reports", force: :cascade do |t|
    t.string "company_name"
    t.string "invoice_no"
    t.string "payment_mode"
    t.datetime "payment_date"
    t.string "y_n"
    t.string "comm"
    t.string "released_before"
    t.string "released_now"
    t.string "on_hold"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end
end

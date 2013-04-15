require 'excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
command.excel_file = File.join(this_directory,'electricity-build-rate-constraint.xlsx')
command.output_directory = this_directory
command.output_name = 'model'
# command.cells_that_can_be_set_at_runtime = { "Sheet1" => ["A1"] }
# command.cells_that_can_be_set_at_runtime = { "Sheet1" => :all }
# command.cells_to_keep = { "Sheet1" => ["A2"]}
command.actually_compile_code = true
command.actually_run_tests = true
command.run_in_memory = true
command.go!

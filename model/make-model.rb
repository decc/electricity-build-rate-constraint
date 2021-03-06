require 'excel_to_code'
root_directory = File.expand_path(File.join(File.dirname(__FILE__), '..'))
command = ExcelToC.new
command.excel_file = File.join(root_directory, 'public', 'electricity-build-rate-constraint.xlsx')
command.output_directory = File.join(root_directory, 'model')
command.output_name = 'model'
# Handy command:
# cut -f 2 electricity-build-rate-constraint/intermediate/Named\ references\ 000 | pbcopy
command.named_references_to_keep = :all
command.named_references_that_can_be_set_at_runtime = :where_possible
command.cells_that_can_be_set_at_runtime = :named_references_only
command.actually_compile_code = true
command.actually_run_tests = true
command.run_in_memory = true
command.go!

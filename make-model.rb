require 'excel_to_code'
this_directory = File.dirname(__FILE__)
command = ExcelToC.new
command.excel_file = File.join(this_directory,'electricity-build-rate-constraint.xlsx')
command.output_directory = this_directory
command.output_name = 'model'
# Handy command:
# cut -f 2 electricity-build-rate-constraint/intermediate/Named\ references\ 000 | pbcopy
command.named_references_to_keep = %w{
Average_life_of_low_carbon_generation
CCS_by_2020
Demand
Electricity_demand_growth_rate
Electricity_demand_in_2012
Electricity_demand_in_2050
Electricity_emissions_during_CB4
Electrification_Start_year
Emissions
Emissions_factor
Emissions_factor_2030
Emissions_factor_2050
High_carbon
High_carbon_EF
High_carbon_emissions_factor_2012
High_carbon_emissions_factor_2020
High_carbon_emissions_factor_2050
High_carbon_load_factor
Low_carbon_load_factor
Maximum_low_c
Maximum_low_carbon_build_rate
Maximum_low_carbon_build_rate_expansion
Maximum_low_carbon_build_rate_expansion
MaxMean2012
MaxMean2050
Minimum_low_carbon_build_rate
MinMean2012
MinMean2050
Net_increase_in_zero_carbon
Nuclear_change_2012_2020
Nuclear_in_2012
Renewable_electricity_in_2020
Renewables_in_2012
Year_second_wave_of_building_starts
Zero_carbon
Zero_carbon_built
Zero_carbon_decomissioned
}
command.named_references_that_can_be_set_at_runtime = %w{
Maximum_low_carbon_build_rate
Year_second_wave_of_building_starts
}
# command.cells_that_can_be_set_at_runtime = { "Sheet1" => ["A1"] }
# command.cells_that_can_be_set_at_runtime = { "Sheet1" => :all }
# command.cells_to_keep = { "Sheet1" => ["A2"]}
command.actually_compile_code = true
command.actually_run_tests = true
command.run_in_memory = true
command.go!

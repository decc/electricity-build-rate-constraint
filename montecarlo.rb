# This runs a montecarlo on the model, dumping the outputs as csv to stdout
# you probably want to redirect them to a file:
#
#     bundle exec ruby montecarlo.rb > public/runs.csv
#
# (c) 2013 Tom Counsell tom@counsell.org
# MIT licence
require './model'


# The inputs to the model that will be varied randomly across the ranges given
# using a flat distribution
parameters = {
  maximum_low_carbon_build_rate: 0..100,
  electrification_start_year: 2020..2040,
  electricity_demand_in_2050: 300..700,
  average_life_of_low_carbon_generation: 20..70,
  renewable_electricity_in_2020: 0.0..0.35,
  ccs_by_2020: 0.0..6.0,
  nuclear_change_2012_2020: -59..10,
  # high_carbon_emissions_factor_2020: 350..650,
  #high_carbon_emissions_factor_2050: 350..650,
  maximum_low_carbon_build_rate_contraction: 0.1..1.0,
  maximum_low_carbon_build_rate_expansion: 0.1..1.0,
  minimum_low_carbon_build_rate: 0.1..10,
  maxmean2050: 1.5..3.0,
  minmean2050: 0.1..0.5,
  annual_change_in_non_electricity_traded_emissions: -0.04..0.01,
}

# The outputs of the model that will be logged with the inputs 
outputs = %w{ 
  electricity_emissions_during_cb4
  emissions_factor_2030
  electricity_emissions_absolute_2050
}

# Output the parameter titles to teh csv
puts (parameters.keys + outputs).join(",")

# The model
m = ModelShim.new

# Use a specific random number generator to take advantage of the #rand(range) method
r = Random.new

# Can vary the number of montecarlo runs
10000.times do
  m.reset
  parameters.each do |parameter,range|
    m.send(parameter.to_s+"=",r.rand(range))
  end
  
  # Outputing the results as csv, one run per line
  puts (parameters.keys.map { |p| m.send(p) } + outputs.map { |o| m.send(o) }).join(",")
end

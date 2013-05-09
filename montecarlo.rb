require './model'

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

outputs = %w{ 
  electricity_emissions_during_cb4
  emissions_factor_2030
  electricity_emissions_absolute_2050
}

puts (parameters.keys + outputs).join(",")

m = ModelShim.new
r = Random.new


10000.times do
  m.reset
  parameters.each do |parameter,range|
    m.send(parameter.to_s+"=",r.rand(range))
  end
  
  puts (parameters.keys.map { |p| m.send(p) } + outputs.map { |o| m.send(o) }).join(",")
end

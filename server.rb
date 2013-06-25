require 'sinatra'
require 'json'
require './model'

# We want to be able to work out which methods are unique to our Model
# and which methods are common to all FFI modules, so we create an empty
# FFI module.
module FFIMethodsToIgnore; extend FFI::Library; end

# This is used to work out what named references exist in the model
def extract_model_structure
  # Get all the excel references
  relevant_methods = (Model.methods - FFIMethodsToIgnore.methods - [:reset])
  # Then remove the ones that look like standard sheet references
  relevant_methods = relevant_methods.find_all do |m|
    m.to_s !~ /^(set_)?model_/
  end
  setters = relevant_methods.find_all do |m|
    m.to_s != /^set_/
  end

  # p setters.map { |m| m[/^set_(.*)$/,1] } # If you want to see the inputs

  # Remove all the setters, because there will be a getter
  relevant_methods = relevant_methods.find_all do |m|
    m.to_s !~ /^set_/
  end
  # And return the array
  relevant_methods
end

# Only work out the model structure once
model_structure = extract_model_structure()

# This should be identical to the one in src/javacripts/chart.js.coffee
url_structure = [
  "version",
  'build_rate_from_now_to_2020',
  'proportion_of_build_rate_to_2020_that_is_wind_rest_is_bio',
  'build_rate_target_in_second_build',
  'proportion_of_second_build_that_is_wind',
  'n_2012_onwards_electricity_demand_growth_rate',
  'year_electricity_demand_starts_to_increase',
  'n_2050_electricity_demand',
  'n_2020_non_renewable_low_carbon_generation_i_e_nuclear_ccs',
  'n_2050_fossil_fuel_emissions_factor',
  'n_2050_maximum_electricity_demand',
  'n_2050_minimum_electricity_demand',
  'annual_change_in_non_electricity_traded_emissions',
  'n_2020_fossil_fuel_emissions_factor',
  'average_life_high_carbon',
  'average_life_other_low_carbon',
  'average_life_wind',
  'maximum_industry_contraction',
  'maximum_industry_expansion',
  'minimum_build_rate'
] - ["version"] # A cludge to make the above lines easier to copy and paste

# This is the method that is used to request data from the model
# the first part of the url is to match a version number the remainder
# should match the url_structure above. 
#
# The method sets the named methods in the url structure to the values passed in the url
# it then alters the year_second_wave_of_building_starts back from 2050 towards 2010 to
# try and get 2050 emissions below 5gCO2/kWh.
#
# It then passes back the results of the model as json
get '/data/1:*' do 
  m = ModelShim.new
  # If a parameter looks like a number, make it a number
  controls =  params[:splat][0].split(':').map { |v| v =~ /^-?\d+\.?\d*$/ ? v.to_f : v }
  year = 2050
  while year >= 2020
    m.reset
    controls.each.with_index do |v,i|
      next unless v && v != ""
      p url_structure[i] + ":" + v.to_s
      m.send(url_structure[i]+"=",v)
    end
    m.year_second_wave_of_building_starts = year
    break if m.n_2050_emissions_electricity < 10
    year = year - 1
  end

  result = {}
  model_structure.each do |method|
    r = m.send(method)
    r.flatten! if r.is_a?(Array) && r.length == 1
    result[method] = r
  end

  result.to_json
end

# The root url. Just returns index.html at the moment
get '*' do
  send_file 'public/index.html'
end

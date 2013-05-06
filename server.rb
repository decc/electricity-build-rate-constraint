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
  p setters.map { |m| m[/^set_(.*)$/,1] }
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
  "maximum_low_carbon_build_rate",
  "electrification_start_year",
  "electricity_demand_in_2050",
  "average_life_of_low_carbon_generation",
  "ccs_by_2020",
  "high_carbon_emissions_factor_2020", 
  "high_carbon_emissions_factor_2050", 
  "maximum_low_carbon_build_rate_contraction", 
  "maximum_low_carbon_build_rate_expansion", 
  "maxmean2050", 
  "minimum_low_carbon_build_rate",
  "minmean2050", 
  "nuclear_change_2012_2020" ,
  "renewable_electricity_in_2020"
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
get '/data/1/*' do 
  m = ModelShim.new
  # If a parameter looks like a number, make it a number
  controls =  params[:splat][0].split('/').map { |v| v =~ /^\d+\.?\d*$/ ? v.to_f : v }
  year = 2050
  while year >= 2010
    m.reset
    controls.each.with_index do |v,i|
      m.send(url_structure[i]+"=",v)
    end
    m.year_second_wave_of_building_starts = year
    break if m.emissions_factor_2050 < 5
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

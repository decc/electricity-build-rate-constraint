require 'sinatra'
require 'json'
require './model'

# We want to be able to work out which methods are unique to our Model
# and which methods are common to all FFI modules, so we create an empty
# FFI module.
module FFIMethodsToIgnore; extend FFI::Library; end

# This is used to work out what named references exist in the model
# and then to assign them to one of three groups:
# inputs - named references that are setable
# series - named references that return an array
# outputs - any other named references
def extract_model_structure
  # Get all the excel references
  relevant_methods = (Model.methods - FFIMethodsToIgnore.methods - [:reset])
  # Then remove the ones that look like standard sheet references
  relevant_methods = relevant_methods.find_all do |m|
    m.to_s !~ /^(set_)?model_/
  end
  structure = { inputs: [], outputs: [], series: [] } 
  relevant_methods.each do |method|
    # If it is a setter, it must be an input
    if method =~ /^set_(.*?)$/
     structure[:inputs] << $1
    # If it is an array, then we add it as a series
    elsif ModelShim.new.send(method).is_a?(Array)
     structure[:series] << method.to_s 
    # Otherwise it must be an output
    else
      structure[:outputs] << method.to_s
    end
  end
  # Remove inputs from outputs
  structure[:outputs] = structure[:outputs] - structure[:inputs]
  structure
end

# Only work out the model structure once
model_structure = extract_model_structure()

# This should be identical to the one in src/javacripts/chart.js.coffee
url_structure = [
  "version",
  "maximum_low_carbon_build_rate",
  "electricity_demand_in_2050"
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
  model_structure.each do |key, value|
    result[key] = h = {}
    value.each do |method|
      r = m.send(method)
      r.flatten! if r.is_a?(Array) && r.length == 1
      h[method] = r
    end
  end

  result.to_json
end

# The root url. Just returns index.html at the moment
get '*' do
  send_file 'public/index.html'
end

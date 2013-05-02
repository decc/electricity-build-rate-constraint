require 'sinatra'
require 'json'
require './model'

# We don't care about standard FFI methods
module FFIMethodsToIgnore; extend FFI::Library; end

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
    # If it is the getter for a setter, we ignore
    elsif structure[:inputs].include?(method.to_s)
     next
    # If it is an array, then we add it as a series
    elsif ModelShim.new.send(method).is_a?(Array)
     structure[:series] << method.to_s 
    # Otherwise it must be an output
    else
      structure[:outputs] << method.to_s
    end
  end
  structure
end

model_structure = extract_model_structure()

get '/data/1/:maximum_low_carbon_build_rate' do 
  m = ModelShim.new
  build_constraint = params[:maximum_low_carbon_build_rate].to_f
  year = 2050
  while year >= 2010
    m.reset
    m.maximum_low_carbon_build_rate = build_constraint
    m.year_second_wave_of_building_starts = year
    break if m.emissions_factor_2050 < 5
    year = year - 1
  end

  p model_structure

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

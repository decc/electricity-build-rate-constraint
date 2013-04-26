require 'sinatra'
require 'json'
require './model'

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

  {
    maximum_low_carbon_build_rate: m.maximum_low_carbon_build_rate,
    year_second_wave_of_building_starts: m.year_second_wave_of_building_starts,
    series: {
      emissions_factor: m.emissions_factor.flatten,
      emissions: m.emissions.flatten,
      zero_carbon_build_rate: m.zero_carbon_built.flatten,
      zero_carbon_output: m.zero_carbon.flatten,
      high_carbon_output: m.high_carbon.flatten,
    }
  }.to_json
end

# The root url. Just returns index.html at the moment
get '*' do
  send_file 'public/index.html'
end


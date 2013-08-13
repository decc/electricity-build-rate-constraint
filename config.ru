# encoding: utf8
require './server'

ENV['RACK_ENV'] = ENV['RAILS_ENV'] if ENV['RAILS_ENV']

if ENV['RAILS_ENV'] == 'development'
  map '/assets' do
    require 'sprockets'
    environment = Sprockets::Environment.new

    environment.append_path 'src/javascripts'
    environment.append_path 'src/stylesheets'
    environment.append_path 'public/assets'
    environment.append_path 'contrib'

    run environment
  end
end

map '/' do
  run Sinatra::Application
end

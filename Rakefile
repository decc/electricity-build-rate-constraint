#!/usr/bin/env rake
# coding: utf-8

require 'rake/clean'

CLEAN.include('model/model.*', 'model/libmodel.dylib', 'model/test_model.rb', 'public/index.html')

task :default => ['model/model.rb']

file 'model/model.rb' => ['public/electricity-build-rate-constraint.xlsx'] do
  require_relative 'model/make-model'
end

require 'sprockets'
require 'rake/sprocketstask'
require 'haml'
require 'json'
require_relative 'src/helper'

# This deals with the javascript and css
environment = Sprockets::Environment.new
environment.append_path 'src/javascripts'
environment.append_path 'src/stylesheets'
environment.append_path 'contrib'

Rake::SprocketsTask.new do |t|
  t.environment = environment
  t.output      = "./public/assets"
  t.assets      = %w( application.js application.css )
end

manifest = './public/assets/manifest.json'
file manifest => ['assets']  

desc "Compiles changes to src/default.html.haml into public/default.html and adds links it to the latest versions of application.cs and application.js"
task 'html' => [manifest] do 

  class Context
    include Helper
  end

  context = Context.new

  # We need to figure out the filename of the latest javascript and css
  context.assets = JSON.parse(IO.readlines(manifest).join)['assets']

  input = IO.readlines('./src/index.html.haml').join
  File.open('./public/index.html','w') do |f|
    f.puts Haml::Engine.new(input).render(context)
  end
end

require 'rake/clean'

CLEAN.include('model/model.*', 'model/libmodel.dylib', 'model/test_model.rb')

task :default => ['model/model.rb']

file 'model/model.rb' => ['public/electricity-build-rate-constraint.xlsx'] do
  require_relative 'model/make-model'
end


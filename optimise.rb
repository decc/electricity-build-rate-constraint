require_relative 'model'

m = ModelShim.new

m.reset
m.model_b9 = rand(100)
puts "#{m.model_b9} TWh/yr/yr so start in #{m.model_b8}"



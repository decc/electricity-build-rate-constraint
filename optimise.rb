require_relative 'model'

m = ModelShim.new

build_constraint = rand(100)
p build_constraint
year = 2050
while year >= 2010
  m.reset
  m.model_b9 = build_constraint
  m.model_b8 = year
  puts "#{m.model_b9} TWh/yr/yr starting in #{m.model_b8} gets to #{m.model_f6} gCO2/kWh in 2050"
  break if m.model_f6 < 5
  year = year - 1
end

puts "Hits target: #{m.model_b9} TWh/yr/yr starting in #{m.model_b8} gets to #{m.model_f6} gCO2/kWh in 2050"


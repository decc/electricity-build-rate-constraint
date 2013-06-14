data = {}
request = undefined
s =
  version: 1

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
]

update = () ->
  if request?
    setUrlToMatchSettings()
    request.abort()
  request = d3.json("/data#{urlForSettings()}", (error, json) ->
    return console.warn(error) if error
    data = json
    window.data = data
    visualise()
  )

setUrlToMatchSettings = () ->
  history.pushState(undefined, undefined, urlForSettings())

window.onpopstate = () ->
  getSettingsFromUrl()
  updateControlsFromSettings()
  update()

urlForSettings = () ->
  url = for a in url_structure
    if s[a]?
      s[a]
    else
      ""
  "/" + url.join(':')

getSettingsFromUrl = () ->
  c = window.location.pathname.substring(1).split(':')

  for a,i in url_structure
    if c[i]? and c[i] != ""
      s[a] = c[i]

timeSeriesChart = ->
  
  margin =
    top: 20
    right: 20
    bottom: 20
    left: 50

  width = 250
  height = 125

  unit = "TWh/yr"
  first_scale_year = 2010
  last_scale_year = 2050
  first_data_year = 2012
  max_value = undefined

  xScale = d3.time.scale()
  yScale = d3.scale.linear()

  xAxis = d3.svg.axis().scale(xScale).orient("bottom").ticks(5)
  yAxis = d3.svg.axis().scale(yScale).orient("left").ticks(5)

  area = d3.svg.area()
    .x((d,i) -> xScale(new Date("#{first_data_year+i}")))
    .y((d,i) -> yScale(+d))

  line = d3.svg.line()
    .x((d,i) -> xScale(new Date("#{first_data_year+i}")))
    .y((d,i) -> yScale(+d))

  chart = (selection) ->
    selection.each (data) ->
      
      # Update the x-scale.
      xScale
        .domain([new Date("#{first_scale_year}"), new Date("#{last_scale_year}")])
        .range([0, width - margin.left - margin.right])
      
      # Update the y-scale.
      yScale
        .domain([0, max_value || d3.max(data)])
        .range([height - margin.top - margin.bottom, 0])
      
      # Select the svg element, if it exists.
      svg = d3.select(this).selectAll("svg").data([data])
      
      # Otherwise, create the skeletal chart.
      gEnter = svg.enter()
        .append("svg")
        .append("g")
        .attr('class','drawing')

      # And the basic bits
      gEnter
        .append("g")
        .attr("class", "x axis")
      gEnter
        .append("g")
        .attr("class", "y axis")
      gEnter
        .append("text")
        .attr("class", "y axislabel")
      
      # Update the outer dimensions.
      svg
        .attr("width", width)
        .attr("height", height)
      
      # Update the inner dimensions.
      g = svg.select("g").attr("transform", "translate(" + margin.left + "," + margin.top + ")")
      
      
      # Update the x-axis.
      g.select(".x.axis")
        .attr("transform", "translate(0," + yScale.range()[0] + ")")
        .call(xAxis)
       
      # Update the y-axis.
      g.select(".y.axis")
        .attr("transform", "translate(0," + xScale.range()[0] + ")")
        .call(yAxis)

      # Update the y-axis label.
      g.select(".y.axislabel")
        .attr("transform", "translate(0," + (xScale.range()[0] - 10) + ")")
        .text(unit)

      # Update the lines
      lines = g.selectAll(".line")
        .data(Object)

      lines.enter()
        .append("path")
        .attr("class", (d,i) -> "line line#{i}")

      lines.transition()
        .attr("d", line)

  chart.unit = (_) ->
    return unit unless _?
    unit = _
    chart

  chart.max_value = (_) ->
    return max_value unless _?
    max_value = _
    chart

  chart

timeSeriesStakedAreaChart = ->
  
  margin =
    top: 20
    right: 20
    bottom: 20
    left: 50

  width = 250
  height = 125

  unit = "TWh/yr"
  first_scale_year = 2010
  last_scale_year = 2050
  first_data_year = 2012
  max_value = undefined

  xScale = d3.time.scale()
  yScale = d3.scale.linear()

  xAxis = d3.svg.axis().scale(xScale).orient("bottom").ticks(5)
  yAxis = d3.svg.axis().scale(yScale).orient("left").ticks(5)

  stack = d3.layout.stack()

  area = d3.svg.area()
    .x((d,i) -> xScale(d.x))
    .y0((d,i) -> yScale(d.y0))
    .y1((d,i) -> yScale(d.y0 + d.y))

  chart = (selection) ->
    selection.each (data) ->

      # Map the data into the right format
      data = data.map(
        (d)-> d.map(
          (p, i) ->
            {x: new Date("#{first_data_year+i}"), y: p }
        )
      )

      layers = stack(data)
      
      # Update the x-scale.
      xScale
        .domain([new Date("#{first_scale_year}"), new Date("#{last_scale_year}")])
        .range([0, width - margin.left - margin.right])
      
      # Update the y-scale.
      yScale
        .domain([0, max_value || d3.max(data)])
        .range([height - margin.top - margin.bottom, 0])
      
      # Select the svg element, if it exists.
      svg = d3.select(this).selectAll("svg").data([layers])
      
      # Otherwise, create the skeletal chart.
      gEnter = svg.enter()
        .append("svg")
        .append("g")
        .attr('class','drawing')

      # And the basic bits
      gEnter
        .append("g")
        .attr("class", "x axis")
      gEnter
        .append("g")
        .attr("class", "y axis")
      gEnter
        .append("text")
        .attr("class", "y axislabel")
      
      # Update the outer dimensions.
      svg
        .attr("width", width)
        .attr("height", height)
      # Update the inner dimensions.
      g = svg.select("g.drawing").attr("transform", "translate(" + margin.left + "," + margin.top + ")")

      # Update the x-axis.
      g.select(".x.axis")
        .attr("transform", "translate(0," + yScale.range()[0] + ")")
        .call(xAxis)
       
      # Update the y-axis.
      g.select(".y.axis")
        .attr("transform", "translate(0," + xScale.range()[0] + ")")
        .call(yAxis)

      # Update the y-axis label.
      g.select(".y.axislabel")
        .attr("transform", "translate(0," + (xScale.range()[0] - 10) + ")")
        .text(unit)

      # Update the area paths
      areas = g.selectAll(".area")
        .data(Object)

      areas.enter()
        .append("path")
        .attr("class", (d,i) -> "area area#{i}")

      areas.transition()
        .attr("d", (d) -> area(d))

  chart.unit = (_) ->
    return unit unless _?
    unit = _
    chart

  chart.max_value = (_) ->
    return max_value unless _?
    max_value = _
    chart

  chart

visualise = () ->

  d3.selectAll('.output')
    .datum(() -> @dataset)
    .text((d) -> d3.format(d.format)(data[d.name]))

  if +data.n_2050_emissions_electricity >= 10
    d3.select('#emissions_warning')
      .transition(1000)
      .style("opacity",1)
  else
    d3.select('#emissions_warning')
      .transition(1000)
      .style("opacity",0)

  d3.select('#build_rates')
    .datum([data.build_rate_total_low_carbon, data.build_rate_high_carbon])
    .call(timeSeriesChart().unit("GW/yr").max_value(12))

  d3.select('#capacity')
    .datum([ data.capacity_total_low_carbon, data.capacity_high_carbon ])
    .call(timeSeriesStakedAreaChart().unit("GW").max_value(150))

  d3.select('#load_factor')
    .datum([data.load_factor_dispatchable_low_carbon, data.load_factor_high_carbon, data.load_factor_demand, data.load_factor_intermittent_low_carbon])
    .call(timeSeriesChart().unit("").max_value(1))

  d3.select('#energy_output')
    .datum([ data.energy_output_total_low_carbon, data.energy_output_high_carbon ])
    .call(timeSeriesStakedAreaChart().unit("TWh/yr").max_value(700))

  d3.select('#emissions_factor')
    .datum([data.emissions_factor])
    .call(timeSeriesChart().unit("gCO2/kWh").max_value(500))

  d3.select('#emissions')
    .datum([data.emissions_electicity, data.emissions_non_electricity_traded])
    .call(timeSeriesStakedAreaChart().unit("MtCO2/yr").max_value(300))

  d3.select('#emissions')
    .datum([data.emissions_uk_share_of_eu_ets_cap_current.slice(0,16), data.emissions_uk_share_of_eu_ets_cap_alternative.slice(0,16)])
    .call(timeSeriesChart().unit("MtCO2/yr").max_value(300))

  # Update the controls to defaults if required
  d3.selectAll('.control')
    .datum(() -> @dataset)
    .filter( (d) -> not s[d.name]? )
    .property('value', (d) -> data[d.name])

d3.selectAll('.control')
  .datum(() -> @dataset)
  .on('change', (d) ->
    # console.log d.name, this.value
    s[d.name] = +this.value
    update()
  )

d3.selectAll('.preset')
  .datum(() -> @dataset)
  .on('click', (d) ->
    s[d.name] = +d.value
    updateControlsFromSettings()
    update()
  )

updateControlsFromSettings = () ->
  d3.selectAll('.control')
    .datum(() -> @dataset)
    .filter( (d) -> s[d.name]? )
    .property('value', (d) -> s[d.name])

getSettingsFromUrl()
updateControlsFromSettings()
update()

data = {}
request = undefined
s =
  version: 1
  maximum_low_carbon_build_rate: 0
  electrification_start_year: 2020
  electricity_demand_in_2050: 600

url_structure = [
  "version",
  "maximum_low_carbon_build_rate",
  "electrification_start_year",
  "electricity_demand_in_2050"
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
    s[a]
  "/" + url.join('/')

getSettingsFromUrl = () ->
  c = window.location.pathname.split('/')
  c.shift()

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
        .append("path")
        .attr("class", "area")
      gEnter
        .append("path")
        .attr("class", "line")
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
      
      # Update the area path.
      g.select(".area")
        .transition()
          .attr("d", area.y0(yScale.range()[0]))
      
      # Update the line path.
      g.select(".line")
        .transition()
          .attr("d", line)
      
      # Update the x-axis.
      g.select(".x.axis")
        .attr("transform", "translate(0," + yScale.range()[0] + ")")
        .call(xAxis)
       
      # Update the y-axis.
      g.select(".y.axis")
        .attr("transform", "translate(0," + xScale.range()[0] + ")")
        .call(yAxis)

      g.select(".y.axislabel")
        .attr("transform", "translate(0," + (xScale.range()[0] - 10) + ")")
        .text(unit)

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
  d3.select('#zero_carbon_build_rate')
    .datum(data.zero_carbon_built)
    .call(timeSeriesChart().unit("TWh/yr/yr").max_value(100))

  d3.select('#emissions')
    .datum(data.emissions)
    .call(timeSeriesChart().unit("MtCO2/yr").max_value(200))

  d3.select('#emissions_factor')
    .datum(data.emissions_factor)
    .call(timeSeriesChart().unit("gCO2/kWh").max_value(500))

  d3.select('#energy_output')
    .datum([ data.zero_carbon, data.high_carbon ])
    .call(timeSeriesStakedAreaChart().unit("TWh/yr").max_value(700))

  d3.selectAll('.output')
    .datum(() -> @dataset)
    .text((d) -> d3.format(d.format)(data[d.name]))

  if +data.emissions_factor_2050 > 5
    d3.select('#emissions_warning')
      .transition(1000)
      .style("opacity",1)
  else
    d3.select('#emissions_warning')
      .transition(1000)
      .style("opacity",0)

d3.selectAll('.control')
  .datum(() -> @dataset)
  .on('change', (d) ->
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
    .property('value', (d) -> s[d.name])

getSettingsFromUrl()
updateControlsFromSettings()
update()
